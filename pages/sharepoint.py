import os
import time
from datetime import datetime, timedelta, timezone
from urllib.parse import parse_qs, unquote, urlparse

import streamlit as st
from office365.runtime.client_request_exception import ClientRequestException
from office365.runtime.queries.service_operation import ServiceOperationQuery
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.recyclebin.item_collection import RecycleBinItemCollection

# SharePoint Online throttles clients that send too many requests, replying with
# HTTP 429 or 503 (usually with a Retry-After header). Bulk runs over large
# libraries reliably trip this, and the library's plain execute_query does not
# retry, so we wrap every call with backoff.
THROTTLE_STATUS_CODES = {429, 503}
MAX_THROTTLE_RETRIES = 5
BASE_BACKOFF_SECONDS = 2

# SharePoint connection details from environment variables.
# Azure Entra ID app-only authentication with a client certificate.
# Note: the legacy ACS (client id + client secret) model was retired by
# Microsoft on 2 April 2026 and no longer works.
TENANT_ID = os.environ.get("TENANT_ID")
CLIENT_ID = os.environ.get("CLIENT_ID")
CERT_THUMBPRINT = os.environ.get("CERT_THUMBPRINT")
CERT_PATH = os.environ.get("CERT_PATH")


def missing_credentials():
    """Return the list of required environment variables that are not set."""
    required = {
        "TENANT_ID": TENANT_ID,
        "CLIENT_ID": CLIENT_ID,
        "CERT_THUMBPRINT": CERT_THUMBPRINT,
        "CERT_PATH": CERT_PATH,
    }
    return [name for name, value in required.items() if not value]


def authenticate_sharepoint(site_url):
    """Authenticate and return the SharePoint client context.

    Uses Azure Entra ID app-only authentication with a client certificate.

    Args:
        site_url: The SharePoint site URL (e.g., https://company.sharepoint.com/sites/SiteName)
    """
    missing = missing_credentials()
    if missing:
        raise ValueError(f"Missing environment variables: {', '.join(missing)}.")
    return ClientContext(site_url).with_client_certificate(
        tenant=TENANT_ID,
        client_id=CLIENT_ID,
        thumbprint=CERT_THUMBPRINT,
        cert_path=CERT_PATH,
    )


def _execute_query_with_retry(ctx):
    """Execute the pending queries, retrying on SharePoint throttling.

    SharePoint Online returns HTTP 429/503 (often with a Retry-After header)
    when a client sends too many requests, which happens on bulk runs over large
    libraries. The library's plain execute_query does not retry, so we honor
    Retry-After when present and otherwise back off exponentially. The failed
    query is re-queued before each retry (the library drops it from the pending
    queue once it starts executing). Non-throttling errors are re-raised at once.
    """
    attempt = 0
    while True:
        try:
            ctx.execute_query()
            return
        except ClientRequestException as error:
            status = getattr(error.response, "status_code", None)
            if status not in THROTTLE_STATUS_CODES or attempt >= MAX_THROTTLE_RETRIES:
                raise

            retry_after = None
            try:
                header = error.response.headers.get("Retry-After")
                if header is not None and str(header).strip().isdigit():
                    retry_after = int(header)
            except (AttributeError, ValueError):
                retry_after = None

            delay = (
                retry_after
                if retry_after is not None
                else BASE_BACKOFF_SECONDS * (2**attempt)
            )
            attempt += 1

            # Re-queue the query that just failed so the retry resends it.
            if ctx.current_query is not None:
                ctx.add_query(ctx.current_query)
            time.sleep(delay)


def extract_site_and_path_from_url(full_url):
    """Extract the SharePoint site URL and server-relative path from a full SharePoint URL.

    Args:
        full_url: Full SharePoint URL, can be:
            - Site only: https://company.sharepoint.com/sites/SiteName
              (also /teams/SiteName, or the root site collection)
            - Library view: https://company.sharepoint.com/sites/SiteName/Shared%20Documents/Forms/AllItems.aspx[?id=...]
            - Direct path: https://company.sharepoint.com/sites/SiteName/Shared%20Documents/folder/file.docx
            - Sharing link with resource path: https://company.sharepoint.com/:f:/r/sites/SiteName/...

    Returns:
        tuple: (site_url, server_relative_path_or_none)
            - site_url: e.g., "https://company.sharepoint.com/sites/SiteName"
            - server_relative_path: e.g., "/sites/SiteName/Shared Documents/folder"
                                   or None if processing entire site
    """
    parsed = urlparse(full_url.strip())
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        raise ValueError("Invalid SharePoint URL: it must start with https://")

    path_parts = [unquote(part) for part in parsed.path.split("/") if part]

    # Sharing links look like /:f:/r/sites/Site/... ("r" = resource path). The
    # tokenized variants (/:f:/s/..., /:f:/g/...) cannot be resolved to a path.
    if path_parts and path_parts[0].startswith(":") and path_parts[0].endswith(":"):
        if len(path_parts) < 2 or path_parts[1] != "r":
            raise ValueError(
                "Unsupported sharing link. Open the item in the browser and copy "
                "the URL from the address bar instead."
            )
        path_parts = path_parts[2:]

    # Site collections live under /sites/<name> or /teams/<name>; anything else
    # is the root site collection.
    if path_parts and path_parts[0] in ("sites", "teams"):
        if len(path_parts) < 2:
            raise ValueError(
                f"Invalid SharePoint URL: site name not found after '{path_parts[0]}'"
            )
        site_parts = path_parts[:2]
        rest = path_parts[2:]
    else:
        site_parts = []
        rest = path_parts

    site_path = "/" + "/".join(site_parts) if site_parts else ""
    site_url = f"{parsed.scheme}://{parsed.netloc}{site_path}"

    # Library views (.../Forms/AllItems.aspx) carry the selected file or folder
    # in the 'id' query parameter.
    query_params = parse_qs(parsed.query)
    if "id" in query_params:
        return site_url, unquote(query_params["id"][0])

    # Without 'id', a library view points at the library root.
    if "Forms" in rest:
        rest = rest[: rest.index("Forms")]
        if not rest:
            raise ValueError("Invalid SharePoint URL: library not found")
        return site_url, f"{site_path}/{'/'.join(rest)}"

    # Nothing after the site, or a site page: process the entire site.
    if not rest or rest[0] in ("SitePages", "_layouts"):
        return site_url, None

    if rest[-1].lower().endswith(".aspx"):
        raise ValueError(
            "Unsupported SharePoint URL: open the file, folder or library in the "
            "browser and copy the URL from the address bar."
        )

    # Direct path to a library, folder or file.
    return site_url, f"{site_path}/{'/'.join(rest)}"


def delete_old_versions_of_file(ctx, file_obj, max_date=None, keep_versions=None):
    """Delete old versions of a single file based on filters.

    Args:
        ctx: SharePoint client context
        file_obj: An already-resolved SharePoint ``File`` object. It MUST come
            from a folder listing (``folder.files``) or ``_resolve_file_by_path``
            so its version requests resolve via ``GetFileById('<guid>')`` rather
            than by re-embedding the server-relative URL. SharePoint returns
            ``401 Unauthorized`` for ``GetFileByServerRelativeUrl`` when the
            decoded path exceeds ~260 characters (a legacy MAX_PATH limit that
            applies to the base URL, not to query parameters), which is exactly
            what happens in deeply nested folders.
        max_date: Delete versions created before this date (datetime object)
        keep_versions: Number of recent versions to keep (int)

    Returns:
        tuple: (versions_deleted_count, message)
    """
    file_url = file_obj.properties.get("ServerRelativeUrl", "<unknown>")
    try:
        versions = file_obj.versions
        ctx.load(versions)
        _execute_query_with_retry(ctx)

        if len(versions) == 0:
            return 0, f"No old versions found for: {file_url}"

        # IMPORTANT: Filter out versions marked as "IsCurrentVersion"
        # SharePoint sometimes includes the current version in the versions collection
        # but won't allow us to delete it
        deletable_versions = [
            v for v in versions if not v.properties.get("IsCurrentVersion", False)
        ]

        if len(deletable_versions) == 0:
            return 0, f"No deletable old versions found for: {file_url}"

        # Filter versions based on criteria
        versions_to_delete = []

        if keep_versions is not None:
            # Keep N versions TOTAL (including current version)
            # Since 'versions' only contains old versions, we need to keep (keep_versions - 1) old versions
            # Example: keep_versions=2 means keep 1 old version + 1 current = 2 total
            old_versions_to_keep = max(0, keep_versions - 1)

            # Sort versions by creation date (most recent first)
            sorted_versions = sorted(
                deletable_versions,
                key=lambda v: v.properties.get("Created", datetime.min),
                reverse=True,
            )
            # Delete all except the N most recent old versions
            versions_to_delete = sorted_versions[old_versions_to_keep:]

        elif max_date is not None:
            # Delete versions created before max_date
            for version in deletable_versions:
                version_date = version.properties.get("Created")
                if version_date and isinstance(version_date, datetime):
                    if version_date < max_date:
                        versions_to_delete.append(version)
                elif version_date and isinstance(version_date, str):
                    # Parse string date if needed
                    # SharePoint dates are UTC; compare as naive UTC like the
                    # datetimes the library parses.
                    try:
                        parsed_date = datetime.fromisoformat(
                            version_date.replace("Z", "+00:00")
                        )
                        if parsed_date.tzinfo is not None:
                            parsed_date = parsed_date.astimezone(timezone.utc).replace(
                                tzinfo=None
                            )
                        if parsed_date < max_date:
                            versions_to_delete.append(version)
                    except (ValueError, AttributeError):
                        pass
        else:
            # No filter: delete ALL versions (original behavior)
            versions_to_delete = list(deletable_versions)

        if len(versions_to_delete) == 0:
            return 0, f"No versions match the deletion criteria for: {file_url}"

        # Delete filtered versions
        for version in versions_to_delete:
            version.delete_object()

        _execute_query_with_retry(ctx)
        return len(
            versions_to_delete
        ), f"Deleted {len(versions_to_delete)} old versions from: {file_url}"

    except Exception as e:
        return -1, f"Error processing {file_url}: {str(e)}"


def get_all_document_libraries(ctx):
    """Get all document libraries in the site, excluding system libraries.

    Returns:
        list: List of tuples (library_title, library_root_folder_url)
    """
    libraries = []

    # System libraries to exclude
    SYSTEM_LIBRARIES = {
        "Form Templates",
        "Site Assets",
        "Style Library",
        "Site Pages",
        "Converted Forms",
        "IWConvertedForms",
        "Master Page Gallery",
        "Theme Gallery",
        "Web Part Gallery",
        "List Template Gallery",
        "Solution Gallery",
    }

    try:
        # Get all lists
        lists = ctx.web.lists
        ctx.load(lists)
        _execute_query_with_retry(ctx)

        for lst in lists:
            # Filter for document libraries: BaseTemplate = 101
            # and not hidden, and not in system libraries
            if (
                lst.properties.get("BaseTemplate") == 101
                and not lst.properties.get("Hidden", False)
                and lst.properties.get("Title") not in SYSTEM_LIBRARIES
            ):
                # Get root folder URL
                root_folder = lst.root_folder
                ctx.load(root_folder)
                _execute_query_with_retry(ctx)

                libraries.append(
                    (lst.properties.get("Title"), root_folder.serverRelativeUrl)
                )

    except Exception as e:
        st.error(f"Error getting document libraries: {str(e)}")

    return libraries


def get_all_files_in_folder(ctx, folder):
    """Recursively get all files in a folder and its subfolders.

    ``folder`` may be a server-relative path (string, used at the top-level entry
    point) or an already-listed ``Folder`` object (used for recursion).

    Returns the ``File`` objects themselves (not their URLs) so that callers can
    operate on them directly. Re-resolving a file by its server-relative URL
    (``GetFileByServerRelativeUrl``) fails with ``401 Unauthorized`` once the
    decoded path exceeds ~260 characters, which is common in deeply nested
    folders. The objects returned here resolve later requests via
    ``GetFileById('<guid>')`` instead, which is immune to path length.

    Recursion MUST pass the ``Folder`` objects from ``folder.folders`` rather
    than their server-relative URLs. Re-resolving a subfolder by URL breaks for
    names containing ``#`` (a URL fragment delimiter): navigating that folder's
    ``.folders`` truncates at the ``#`` and resolves to its PARENT, whose listing
    includes the ``#`` folder again, causing infinite recursion
    (``maximum recursion depth exceeded``). The listed objects navigate via
    ``GetFolderById('<guid>')``, which is immune to this.
    """
    files = []
    folder_label = (
        folder
        if isinstance(folder, str)
        else folder.properties.get("ServerRelativeUrl", "<folder>")
    )

    try:
        # Resolve the top-level folder from its path once; recursion receives the
        # Folder object directly.
        if isinstance(folder, str):
            folder = ctx.web.get_folder_by_server_relative_url(folder)
            ctx.load(folder)
            _execute_query_with_retry(ctx)

        # Get files in current folder
        folder_files = folder.files
        ctx.load(folder_files)
        _execute_query_with_retry(ctx)

        files.extend(folder_files)

        # Get subfolders and process recursively (passing the objects, not URLs)
        subfolders = folder.folders
        ctx.load(subfolders)
        _execute_query_with_retry(ctx)

        for subfolder in subfolders:
            # Skip system folders
            if not subfolder.name.startswith("."):
                files.extend(get_all_files_in_folder(ctx, subfolder))

    except Exception as e:
        st.warning(f"Could not access folder {folder_label}: {str(e)}")

    return files


def _resolve_file_by_path(ctx, path):
    """Return a ``File`` object for a server-relative path, or None if not found.

    Tries the direct ``GetFileByServerRelativeUrl`` lookup first. That call
    returns ``401 Unauthorized`` when the decoded path exceeds ~260 characters
    (a legacy MAX_PATH limit that applies to the base URL). In that case we fall
    back to listing the parent folder and matching the file by name; the object
    from the listing resolves later requests via ``GetFileById('<guid>')``,
    which is immune to path length.
    """
    try:
        file_obj = ctx.web.get_file_by_server_relative_url(path)
        ctx.load(file_obj)
        _execute_query_with_retry(ctx)
        return file_obj
    except Exception:
        parent, _, name = path.rpartition("/")
        if not parent or not name:
            return None
        try:
            folder = ctx.web.get_folder_by_server_relative_url(parent)
            folder_files = folder.files
            ctx.load(folder_files)
            _execute_query_with_retry(ctx)
            for candidate in folder_files:
                if candidate.properties.get("Name") == name:
                    return candidate
        except Exception:
            return None
        return None


def process_entire_site(ctx, max_date=None, keep_versions=None):
    """Process all document libraries in the entire site.

    Args:
        ctx: SharePoint client context
        max_date: Delete versions created before this date (datetime object)
        keep_versions: Number of recent versions to keep (int)

    Returns:
        dict: Processing results with statistics
    """
    results = {
        "total_files": 0,
        "processed_files": 0,
        "total_versions_deleted": 0,
        "total_libraries": 0,
        "processed_libraries": 0,
        "errors": [],
        "successes": [],
    }

    # Get all document libraries
    st.info("🔍 Discovering document libraries in the site...")
    libraries = get_all_document_libraries(ctx)
    results["total_libraries"] = len(libraries)

    if not libraries:
        st.warning("No document libraries found in this site.")
        return results

    st.success(f"✅ Found {len(libraries)} document libraries")
    st.info("📚 Libraries to process:")
    for lib_title, _ in libraries:
        st.write(f"  • {lib_title}")

    st.divider()

    # Process each library
    for lib_index, (lib_title, lib_root_url) in enumerate(libraries, 1):
        st.subheader(f"📁 Library {lib_index}/{len(libraries)}: {lib_title}")

        try:
            # Get all files in this library
            st.info(f"Scanning {lib_title}...")
            files = get_all_files_in_folder(ctx, lib_root_url)

            if not files:
                st.info(f"No files found in {lib_title}")
                results["processed_libraries"] += 1
                continue

            st.success(f"Found {len(files)} files in {lib_title}")
            results["total_files"] += len(files)

            # Process files with progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file_obj in enumerate(files):
                file_name = file_obj.properties.get("Name", "")
                status_text.text(f"Processing: {file_name} ({i + 1}/{len(files)})")
                progress_bar.progress((i + 1) / len(files))

                versions_deleted, message = delete_old_versions_of_file(
                    ctx, file_obj, max_date, keep_versions
                )

                if versions_deleted >= 0:
                    results["processed_files"] += 1
                    results["total_versions_deleted"] += versions_deleted
                    if versions_deleted > 0:
                        results["successes"].append(message)
                else:
                    results["errors"].append(message)

            progress_bar.empty()
            status_text.empty()

            results["processed_libraries"] += 1
            st.success(f"✅ Completed {lib_title}")

        except Exception as e:
            st.error(f"❌ Error processing library {lib_title}: {str(e)}")
            results["errors"].append(f"Library '{lib_title}': {str(e)}")

        st.divider()

    return results


def process_path(ctx, path, max_date=None, keep_versions=None):
    """Process either a single file or all files in a folder recursively."""
    results = {
        "total_files": 0,
        "processed_files": 0,
        "total_versions_deleted": 0,
        "errors": [],
        "successes": [],
    }

    # A path is a file only if SharePoint resolves it as one. Guessing from the
    # name is unreliable: folders can contain dots ("15.SUBMISSION") and files
    # can lack an extension.
    file_obj = _resolve_file_by_path(ctx, path)

    if file_obj is not None:
        # Single file
        st.info(f"📄 Processing single file: {path}")
        results["total_files"] = 1
        versions_deleted, message = delete_old_versions_of_file(
            ctx, file_obj, max_date, keep_versions
        )

        if versions_deleted >= 0:
            results["processed_files"] = 1
            results["total_versions_deleted"] = versions_deleted
            results["successes"].append(message)
        else:
            results["errors"].append(message)
    else:
        # Folder - process all files recursively
        st.info(f"📁 Processing folder recursively: {path}")
        files = get_all_files_in_folder(ctx, path)
        results["total_files"] = len(files)

        if files:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file_obj in enumerate(files):
                file_name = file_obj.properties.get("Name", "")
                status_text.text(f"Processing: {file_name}")
                progress_bar.progress((i + 1) / len(files))

                versions_deleted, message = delete_old_versions_of_file(
                    ctx, file_obj, max_date, keep_versions
                )

                if versions_deleted >= 0:
                    results["processed_files"] += 1
                    results["total_versions_deleted"] += versions_deleted
                    if versions_deleted > 0:
                        results["successes"].append(message)
                else:
                    results["errors"].append(message)

            progress_bar.empty()
            status_text.empty()
        else:
            st.warning("No files found in the specified folder.")

    return results


def _parse_deleted_date(item):
    """Return a recycle bin item's DeletedDate as a UTC-aware datetime or None.

    Args:
        item: A RecycleBinItem instance.

    Returns:
        datetime or None: Timezone-aware UTC datetime, or None if unparseable.
    """
    value = item.properties.get("DeletedDate")

    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, str):
        try:
            parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except (ValueError, AttributeError):
            return None
    else:
        return None

    # SharePoint returns DeletedDate in UTC; normalise naive values to UTC.
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)

    return parsed


def _load_first_stage_recycle_items(ctx, page_size=2000):
    """Return every first-stage item of the site-collection recycle bin, paged.

    Enumeration constraints we worked around (verified against SharePoint Online):

    - Under app-only certificate auth the WEB recycle bin
      (``ctx.web.recycle_bin``) returns 0 items, so we query the SITE COLLECTION
      recycle bin (``ctx.site`` ``GetRecycleBinItems``), which returns every
      first-stage item regardless of who deleted it.
    - A single unbounded ``ctx.load`` of the whole collection risks failing once
      a bin grows past the server's list-view limits, so we page instead.
    - Office365-REST-Python-Client 2.6.2 exposes no working ``pagingInfo``
      continuation token for the recycle bin (every format the server rejects
      with "searchPos out of range"). What DOES work reliably: an ascending
      ``rowLimit`` query returns the first N items as a stable prefix (verified:
      each page is a strict prefix of a larger request, no duplicates). So we
      request growing prefixes until a response comes back short, which means we
      reached the end of the bin.

    Args:
        ctx: Authenticated SharePoint client context for the site.
        page_size: How many extra items to request on each round.

    Returns:
        list: RecycleBinItem objects with ItemState == 1 (first stage).
    """
    site = ctx.site
    row_limit = page_size
    items = []
    while True:
        result = RecycleBinItemCollection(ctx, site.recycle_bin.resource_path)
        payload = {
            "rowLimit": row_limit,
            "isAscending": True,
            "pagingInfo": None,
            "itemState": 1,
        }
        qry = ServiceOperationQuery(
            site, "GetRecycleBinItems", None, payload, None, result
        )
        ctx.add_query(qry)
        _execute_query_with_retry(ctx)

        items = list(result)
        # A short page (fewer items than requested) means the ascending prefix
        # already covers the whole bin; anything else means there may be more.
        if len(items) < row_limit:
            break
        row_limit += page_size

    return items


def move_recycle_bin_to_second_stage(ctx, cutoff_date=None, run_start=None):
    """Move first-stage recycle bin items to the second-stage recycle bin.

    The second-stage (site collection) recycle bin does not count against the
    site storage quota, so moving items there frees space at the site level.
    Items stay recoverable for the rest of the 93-day retention window (see the
    verification note below).

    Note: we read the recycle bin from the SITE COLLECTION (via
    ``_load_first_stage_recycle_items``), not the web
    (``ctx.web.recycle_bin``). Under app-only certificate authentication the
    web-level recycle bin endpoint returns 0 items (it is filtered to items
    deleted by the calling principal, and version deletions usually do
    not go to the bin), whereas the site-collection endpoint returns every
    first-stage item regardless of who deleted it. Using the web-level bin
    silently reported "already empty" even when the site recycle bin was full of
    user-deleted files. The read is paged so it holds up on large bins.

    Note: moving to the second stage is NOT a permanent delete here. Verified
    against SharePoint Online and Microsoft docs: moved items keep ItemState == 2 and
    stay recoverable for the remainder of the 93-day retention window. (Only
    ``delete_object`` / emptying the second stage purges permanently.)

    Note: the collection-level ``move_to_second_stage_by_ids`` is broken in
    Office365-REST-Python-Client 2.6.2 (it sends a malformed body and the
    server returns HTTP 400), so we move each item individually with the
    per-item ``move_to_second_stage`` method, which works.

    Args:
        ctx: Authenticated SharePoint client context for the site.
        cutoff_date: If None, move every first-stage item. If a UTC-aware
            datetime, move only items deleted before it.
        run_start: UTC-aware datetime marking the start of this run. When a
            cutoff_date is set, items deleted at or after run_start (the
            versions this tool just removed) are moved too, regardless of the
            date filter.

    Returns:
        tuple: (moved_count, message). moved_count is -1 on error.
    """
    try:
        # Enumerate the first-stage site-collection recycle bin, paged.
        first_stage = _load_first_stage_recycle_items(ctx)

        if not first_stage:
            return 0, "First-stage recycle bin is already empty."

        if cutoff_date is None:
            items_to_move = first_stage
        else:
            items_to_move = []
            for item in first_stage:
                deleted_at = _parse_deleted_date(item)
                if deleted_at is None:
                    continue
                is_old = deleted_at < cutoff_date
                is_from_this_run = run_start is not None and deleted_at >= run_start
                if is_old or is_from_this_run:
                    items_to_move.append(item)

        if not items_to_move:
            return 0, "No recycle bin items match the move criteria."

        # Move each item individually, flushing in batches to keep each
        # request reasonable.
        moved = 0
        batch = 0
        for item in items_to_move:
            item.move_to_second_stage()
            moved += 1
            batch += 1
            if batch >= 200:
                _execute_query_with_retry(ctx)
                batch = 0
        if batch:
            _execute_query_with_retry(ctx)

        return moved, f"Moved {moved} item(s) to the second-stage recycle bin."

    except Exception as e:
        return -1, f"Error moving recycle bin items to second stage: {str(e)}"


# Initialize session state variables if they don't exist
if "input_url" not in st.session_state:
    st.session_state.input_url = ""


# Function to handle form submission
def handle_clean_button():
    if not st.session_state.input_url.strip():
        st.error("Please enter at least one SharePoint URL.")
        return

    # Parse multiple URLs (one per line)
    urls_input = st.session_state.input_url.strip()
    url_list = [url.strip() for url in urls_input.split("\n") if url.strip()]

    if not url_list:
        st.error("Please enter at least one valid SharePoint URL.")
        return

    st.info(f"📋 Processing {len(url_list)} URL(s)")

    # Timestamp marking the start of this run (UTC). Used to always move the
    # versions removed in this run out of the first-stage recycle bin, even
    # when a date filter is set.
    run_start = datetime.now(timezone.utc)

    # Get filter values from session state
    filter_type = st.session_state.get("filter_type", "all")
    max_date = None
    keep_versions = None

    # Validate and set filter parameters
    if filter_type == "date":
        if "max_date_filter" in st.session_state and st.session_state.max_date_filter:
            max_date = datetime.combine(
                st.session_state.max_date_filter, datetime.min.time()
            )
    elif filter_type == "versions":
        if (
            "keep_versions_filter" in st.session_state
            and st.session_state.keep_versions_filter
        ):
            keep_versions = st.session_state.keep_versions_filter

    # Display active filter
    if filter_type == "date" and max_date:
        st.info(
            f"🗓️ Filter: Deleting versions created before {max_date.strftime('%Y-%m-%d')}"
        )
    elif filter_type == "versions" and keep_versions:
        st.info(
            f"🔢 Filter: Keeping {keep_versions} total versions (current + {max(0, keep_versions - 1)} old versions)"
        )
    else:
        st.info("🗑️ Filter: Deleting ALL old versions")

    st.divider()

    # Aggregate results across all URLs
    total_results = {
        "total_files": 0,
        "processed_files": 0,
        "total_versions_deleted": 0,
        "total_libraries": 0,
        "processed_libraries": 0,
        "errors": [],
        "successes": [],
        "urls_processed": 0,
        "urls_failed": 0,
    }

    # Reuse one authenticated context per site across its URLs.
    site_contexts = {}

    # Process each URL
    for url_index, url in enumerate(url_list, 1):
        st.subheader(f"🔗 URL {url_index}/{len(url_list)}")
        st.write(
            f"Processing: `{url[:80]}...`" if len(url) > 80 else f"Processing: `{url}`"
        )

        try:
            with st.spinner(f"Processing URL {url_index}..."):
                # Extract site URL and path from the provided URL
                site_url, path = extract_site_and_path_from_url(url)
                st.info(f"🔗 Connecting to: {site_url}")

                # Authenticate once per site and reuse across its URLs.
                if site_url in site_contexts:
                    ctx = site_contexts[site_url]
                else:
                    ctx = authenticate_sharepoint(site_url)
                    site_contexts[site_url] = ctx

                # Determine processing scope
                if path is None:
                    # Entire site processing
                    st.warning(
                        "🌐 **ENTIRE SITE MODE**: Processing ALL document libraries in this site!"
                    )
                    st.info(
                        "This may take a significant amount of time depending on the site size."
                    )
                    results = process_entire_site(ctx, max_date, keep_versions)
                else:
                    # Process specific path (file or folder)
                    results = process_path(ctx, path, max_date, keep_versions)

                # Aggregate results
                total_results["total_files"] += results.get("total_files", 0)
                total_results["processed_files"] += results.get("processed_files", 0)
                total_results["total_versions_deleted"] += results.get(
                    "total_versions_deleted", 0
                )
                total_results["total_libraries"] += results.get("total_libraries", 0)
                total_results["processed_libraries"] += results.get(
                    "processed_libraries", 0
                )
                total_results["errors"].extend(results.get("errors", []))
                total_results["successes"].extend(results.get("successes", []))
                total_results["urls_processed"] += 1

                st.success(f"✅ URL {url_index} completed successfully!")

        except ValueError as e:
            st.error(f"❌ URL {url_index} Error: {str(e)}")
            total_results["errors"].append(f"URL {url_index} ({url}): {str(e)}")
            total_results["urls_failed"] += 1
        except Exception as e:
            st.error(f"❌ URL {url_index} Error: {str(e)}")
            total_results["errors"].append(f"URL {url_index} ({url}): {str(e)}")
            total_results["urls_failed"] += 1

        st.divider()

    # Optionally move each processed site's recycle bin to the second stage.
    recycle_moved_total = 0
    if st.session_state.get("clean_recycle_bin") and site_contexts:
        st.subheader("🗑️ Moving recycle bin to the second stage")

        cutoff = None
        if st.session_state.get("filter_recycle_by_date") and st.session_state.get(
            "recycle_cutoff_date"
        ):
            chosen = st.session_state.recycle_cutoff_date
            cutoff = datetime(
                chosen.year, chosen.month, chosen.day, tzinfo=timezone.utc
            )
            st.info(
                f"🗓️ Moving items deleted before {chosen.isoformat()} "
                "(plus the versions removed in this run)."
            )
        else:
            st.info("🗑️ Moving ALL first-stage recycle bin items.")

        for site_url, ctx in site_contexts.items():
            with st.spinner(f"Processing recycle bin: {site_url}"):
                moved, message = move_recycle_bin_to_second_stage(
                    ctx, cutoff_date=cutoff, run_start=run_start
                )
            if moved >= 0:
                recycle_moved_total += moved
                st.success(f"✅ {site_url}: {message}")
            else:
                st.error(f"❌ {site_url}: {message}")
                total_results["errors"].append(f"Recycle bin ({site_url}): {message}")

        st.divider()

    # Display final aggregate results
    st.success("✅ **ALL URLs PROCESSED!**")
    st.subheader("📊 Summary Statistics")

    col1, col2 = st.columns(2)
    with col1:
        st.metric("URLs Processed", total_results["urls_processed"])
        st.metric("URLs Failed", total_results["urls_failed"])
    with col2:
        st.metric("Total Files", total_results["total_files"])
        st.metric("Versions Deleted", total_results["total_versions_deleted"])

    if total_results["total_libraries"] > 0:
        st.write(f"**Document libraries found:** {total_results['total_libraries']}")
        st.write(
            f"**Document libraries processed:** {total_results['processed_libraries']}"
        )

    st.write(f"**Files processed:** {total_results['processed_files']}")

    if st.session_state.get("clean_recycle_bin"):
        st.write(f"**Recycle bin items moved to second stage:** {recycle_moved_total}")

    if total_results["errors"]:
        st.error(f"**Total errors encountered:** {len(total_results['errors'])}")
        with st.expander("View all errors"):
            for error in total_results["errors"]:
                st.write(f"❌ {error}")

    if total_results["successes"]:
        with st.expander(
            f"View all successful operations ({len(total_results['successes'])})"
        ):
            for success in total_results["successes"]:
                st.write(f"✅ {success}")

    # Reset form field
    st.session_state.input_url = ""


# Main UI
st.title("🧹 SharePoint Version Cleaner")

st.markdown(
    "This tool deletes old versions of SharePoint files based on your "
    "filtering criteria.\n\n⚠️ This action **cannot be undone**"
)

# Check if required environment variables are set
_missing = missing_credentials()
if _missing:
    st.error(
        f"❌ Missing required environment variables ({', '.join(_missing)}). Please check your .env configuration."
    )
    st.stop()

urls = st.text_area(
    "SharePoint URLs (one per line)",
    placeholder="Paste SharePoint URLs here (site, file, or folder)...\nOne URL per line for batch processing",
    key="input_url",
    help="Copy URLs from your browser. You can paste multiple URLs, one per line, to process them all at once.",
    height=150,
)

# Filtering options
st.subheader("🔍 Filtering Options")

filter_type = st.radio(
    "Select deletion filter:",
    options=["all", "date", "versions"],
    format_func=lambda x: {
        "all": "🗑️ Delete ALL old versions (keep only current)",
        "date": "📅 Delete versions older than a specific date",
        "versions": "🔢 Keep a specific number of total versions (including current)",
    }[x],
    key="filter_type",
    help="Choose how to filter which versions to delete",
)

# Show additional inputs based on filter type
if filter_type == "date":
    st.date_input(
        "Delete versions created before:",
        value=None,
        key="max_date_filter",
        help="Only versions created before this date will be deleted",
        format="YYYY-MM-DD",
    )
    st.caption(
        "💡 Example: If you select 2024-01-01, all versions created before January 1st, 2024 will be deleted"
    )

elif filter_type == "versions":
    st.number_input(
        "Total number of versions to keep (including current):",
        min_value=1,
        max_value=100,
        value=5,
        step=1,
        key="keep_versions_filter",
        help="Keep this many total versions (including the current version), delete all older versions",
    )
    st.caption(
        "💡 Example: If you enter 2, only the current version + 1 old version will be kept. If you enter 1, only the current version will be kept (all history deleted)"
    )

st.divider()

# Recycle bin options
st.subheader("🗑️ Recycle Bin")

st.checkbox(
    "Move the site recycle bin to the second stage after cleaning",
    key="clean_recycle_bin",
    help=(
        "The second-stage (site collection) recycle bin does not count against "
        "the site storage quota, so moving items there frees space while they "
        "stay recoverable for the rest of the 93-day retention window."
    ),
)

if st.session_state.get("clean_recycle_bin"):
    st.checkbox(
        "Only move items deleted before a specific date",
        key="filter_recycle_by_date",
        help=(
            "When disabled, every first-stage item is moved. When enabled, only "
            "items deleted before the chosen date are moved."
        ),
    )

    if st.session_state.get("filter_recycle_by_date"):
        default_cutoff = (datetime.now(timezone.utc) - timedelta(days=30)).date()
        st.date_input(
            "Move items deleted before:",
            value=default_cutoff,
            key="recycle_cutoff_date",
            format="YYYY-MM-DD",
        )
        st.caption(
            "💡 Items deleted before this date move to the second stage. More "
            "recent deletions stay in the first-stage bin, except the versions "
            "this run just removed, which are always moved."
        )

st.divider()

col1, col2 = st.columns([1, 4])
with col1:
    st.button("Clean", on_click=handle_clean_button, type="primary")

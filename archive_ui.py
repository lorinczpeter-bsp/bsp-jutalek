"""Drive connection adapter for background persistence."""
from drive_archive import ArchiveError, DriveArchive


def archive_client(st):
    try:
        config = dict(st.secrets.get('drive_archive', {}))
        credentials = dict(st.secrets.get('gcp_service_account', {}))
    except Exception:
        config, credentials = {}, {}
    if not config.get('folder_id') or not credentials:
        return None
    return DriveArchive(config, credentials)


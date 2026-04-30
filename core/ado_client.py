import os
import tempfile
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import requests
import urllib3

from config import ADO_ORG, ADO_PROJECT, ADO_PIPELINE_ID, ADO_ARTIFACT_NAME, ADO_ARTIFACT_FILE

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_BASE = f"https://dev.azure.com/{ADO_ORG}/{ADO_PROJECT}"
_API = "7.1"
_SSL = False  # Corporate proxy intercepts TLS; no local CA bundle available


def _auth(pat):
    return ("", pat)


def fetch_builds(pat, top=20):
    url = f"{_BASE}/_apis/build/builds"
    params = {
        "definitions": ADO_PIPELINE_ID,
        "$top": top,
        "queryOrder": "finishTimeDescending",
        "api-version": _API,
    }
    try:
        r = requests.get(url, auth=_auth(pat), params=params, timeout=15, verify=_SSL)
        r.raise_for_status()
        return True, r.json().get("value", [])
    except requests.HTTPError as e:
        return False, f"HTTP {e.response.status_code}: {e.response.reason}"
    except requests.RequestException as e:
        return False, str(e)


def _file_download_url(download_url):
    parsed = urlparse(download_url)
    params = parse_qs(parsed.query, keep_blank_values=True)
    params["format"] = ["file"]
    params["subPath"] = [f"/{ADO_ARTIFACT_FILE}"]
    query = urlencode({k: v[0] for k, v in params.items()})
    return urlunparse(parsed._replace(query=query))


def download_artifact_zip(pat, build_id, progress_cb=None):
    url = f"{_BASE}/_apis/build/builds/{build_id}/artifacts"
    params = {"artifactName": ADO_ARTIFACT_NAME, "api-version": _API}
    try:
        r = requests.get(url, auth=_auth(pat), params=params, timeout=15, verify=_SSL)
        r.raise_for_status()
        download_url = r.json()["resource"]["downloadUrl"]
    except (requests.RequestException, KeyError) as e:
        return False, f"Could not get artifact URL: {e}"

    file_url = _file_download_url(download_url)
    dest = os.path.join(tempfile.gettempdir(), f"foxl_deploy_{build_id}.zip")

    try:
        with requests.get(file_url, auth=_auth(pat), stream=True, timeout=120, verify=_SSL) as r:
            r.raise_for_status()
            total = int(r.headers.get("content-length", 0))
            downloaded = 0
            last_pct = -1
            with open(dest, "wb") as f:
                for chunk in r.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_cb and total:
                            pct = int(downloaded / total * 100)
                            if pct != last_pct:
                                last_pct = pct
                                progress_cb(downloaded, total)
        return True, dest
    except requests.RequestException as e:
        return False, f"Download failed: {e}"

import hashlib
import os
import re
import time
import pandas as pd
import requests

# ============================================================
# PENGATURAN
# ============================================================
METODE = "apify"                       # "apify" atau "research_api"

LINK_VIDEO = [                         # boleh lebih dari satu video
    "https://vt.tiktok.com/ZSpUANq2n/",
]
MAKS_KOMENTAR = 500                    # maksimal komentar per video

APIFY_TOKEN = os.environ.get("APIFY_TOKEN", "isi personal api key Anda di sini")
APIFY_ACTOR = "fetch_cat/tiktok-comments-scraper"   # bisa diganti actor lain

TIKTOK_TOKEN = os.environ.get("TIKTOK_TOKEN", "isi personal api key Anda di sini")

FILE_EXCEL = "dataset_tiktok.xlsx"
FILE_CSV = "dataset_tiktok.csv"

ANONIMKAN_USERNAME = True              # nama pengguna di-hash (sesuai etika di laporan)


# ============================================================
# BANTUAN
# ============================================================
def url_lengkap(link):
    """Short link (vt.tiktok.com) -> URL video lengkap dengan /video/<id>."""
    if "/video/" in link:
        return link.split("?")[0]
    resp = requests.get(link, allow_redirects=True, timeout=30,
                        headers={"User-Agent": "Mozilla/5.0"})
    url = resp.url.split("?")[0]
    if not re.search(r"/video/\d+", url):
        raise ValueError(f"ID video tidak ditemukan untuk link: {link}")
    return url


def id_video(url):
    return int(re.search(r"/video/(\d+)", url).group(1))


def ambil(item, *kunci, default=None):
    """Ambil nilai dari dict dengan beberapa kemungkinan nama field."""
    for k in kunci:
        if k in item and item[k] is not None:
            return item[k]
    return default


def anonim(nama):
    if not nama:
        return None
    return "user_" + hashlib.sha256(str(nama).encode("utf-8")).hexdigest()[:8]


# ============================================================
# METODE 1: APIFY
# ============================================================
def scrape_apify(urls, maks):
    try:
        from apify_client import ApifyClient
    except ImportError:
        raise SystemExit("Jalankan dulu: pip install apify-client")

    client = ApifyClient(APIFY_TOKEN)
    run = client.actor(APIFY_ACTOR).call(run_input={
        "videoUrls": urls,
        "maxCommentsPerVideo": maks,
        "includeReplies": False,
    })
    dataset_id = run.default_dataset_id if hasattr(run, "default_dataset_id") else run["defaultDatasetId"]
    return list(client.dataset(dataset_id).iterate_items())


def rapikan_apify(items):
    baris = []
    for it in items:
        baris.append({
            "ID Komentar": ambil(it, "commentId", "cid", "id"),
            "ID Video": ambil(it, "videoId"),
            "Komentar": ambil(it, "text", "comment", "commentText"),
            "Username": ambil(it, "authorUsername", "uniqueId", "username"),
            "Like": ambil(it, "likeCount", "diggCount", "likes", default=0),
            "Jumlah Balasan": ambil(it, "replyCount", "replyCommentTotal", default=0),
            "Waktu": ambil(it, "createdAt", "createTimeISO", "createTime"),
            "Link Video": ambil(it, "videoUrl"),
        })
    return pd.DataFrame(baris)


# ============================================================
# METODE 2: TikTok Research API
# ============================================================
def scrape_research_api(urls, maks):
    api = "https://open.tiktokapis.com/v2/research/video/comment/list/"
    headers = {"Authorization": f"Bearer {TIKTOK_TOKEN}", "Content-Type": "application/json"}
    params = {"fields": "id,video_id,text,like_count,reply_count,create_time"}

    semua = []
    for url in urls:
        vid, cursor, jumlah = id_video(url), 0, 0
        while jumlah < maks:
            body = {"video_id": vid, "max_count": 100, "cursor": cursor}
            hasil = requests.post(api, headers=headers, params=params,
                                  json=body, timeout=30).json()
            if hasil.get("error", {}).get("code") != "ok":
                print("Error:", hasil.get("error"))
                break
            data = hasil.get("data", {})
            for k in data.get("comments", []):
                k["videoUrl"] = url
                semua.append(k)
                jumlah += 1
            if not data.get("has_more"):
                break
            cursor = data.get("cursor")
            time.sleep(1)
    return semua


def rapikan_research(items):
    baris = []
    for it in items:
        baris.append({
            "ID Komentar": it.get("id"),
            "ID Video": it.get("video_id"),
            "Komentar": it.get("text"),
            "Username": None,   # Research API tidak memberi username di endpoint ini
            "Like": it.get("like_count", 0),
            "Jumlah Balasan": it.get("reply_count", 0),
            "Waktu": pd.to_datetime(it.get("create_time"), unit="s", errors="coerce"),
            "Link Video": it.get("videoUrl"),
        })
    return pd.DataFrame(baris)


# ============================================================
# PEMBERSIHAN DATASET & EKSPOR
# ============================================================
def finalisasi(df):
    df = df.dropna(subset=["Komentar"]).copy()
    df["Komentar"] = df["Komentar"].astype(str).str.strip()
    df = df[df["Komentar"] != ""]
    df = df.drop_duplicates(subset=["ID Video", "ID Komentar"]).reset_index(drop=True)

    if ANONIMKAN_USERNAME:
        df["Username"] = df["Username"].apply(anonim)

    df.insert(0, "No", range(1, len(df) + 1))
    return df


def simpan(df):
    df.to_excel(FILE_EXCEL, index=False)
    # Excel -> CSV (utf-8-sig supaya emoji/huruf terbaca benar saat dibuka di Excel)
    pd.read_excel(FILE_EXCEL).to_csv(FILE_CSV, index=False, encoding="utf-8-sig")
    print(f"Tersimpan: {FILE_EXCEL} dan {FILE_CSV}")


def jalankan():
    urls = [url_lengkap(l) for l in LINK_VIDEO]
    print("Video:", *urls, sep="\n  ")

    if METODE == "apify":
        df = rapikan_apify(scrape_apify(urls, MAKS_KOMENTAR))
    elif METODE == "research_api":
        df = rapikan_research(scrape_research_api(urls, MAKS_KOMENTAR))
    else:
        raise ValueError("METODE harus 'apify' atau 'research_api'")

    if df.empty:
        raise SystemExit("Tidak ada komentar yang terambil. Cek link, token, atau metode.")

    df = finalisasi(df)
    simpan(df)
    print("Jumlah komentar:", len(df))
    print(df[["No", "Komentar", "Like", "Waktu"]].head())
    return df


if __name__ == "__main__":
    jalankan()
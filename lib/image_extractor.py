"""記事URLから代表画像URLを抽出するモジュール"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from .config import HTTP_HEADERS, HTTP_TIMEOUT


def extract_image_url(article_url: str) -> str | None:
    """記事URLからog:image等の代表画像URLを抽出する

    優先順位:
    1. og:image メタタグ
    2. twitter:image メタタグ
    3. article/main内の最初の大きい画像

    Args:
        article_url: 記事のURL

    Returns:
        画像のURL、取得できない場合はNone
    """
    if not article_url:
        return None

    try:
        response = requests.get(
            article_url, headers=HTTP_HEADERS, timeout=HTTP_TIMEOUT
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")

        # Priority 1: og:image
        og_image = soup.find("meta", property="og:image")
        if og_image and og_image.get("content"):
            img_url = og_image["content"].strip()
            if img_url:
                return _make_absolute(img_url, article_url)

        # Priority 2: twitter:image
        tw_image = soup.find("meta", attrs={"name": "twitter:image"})
        if tw_image and tw_image.get("content"):
            img_url = tw_image["content"].strip()
            if img_url:
                return _make_absolute(img_url, article_url)

        # Priority 3: 記事本文内の最初の意味のある画像
        article_tag = soup.find("article") or soup.find("main") or soup.find("body")
        if article_tag:
            for img in article_tag.find_all("img"):
                src = img.get("src") or img.get("data-src") or img.get("data-lazy-src")
                if not src:
                    continue

                # 小さい画像をスキップ
                width = img.get("width")
                height = img.get("height")
                if width:
                    try:
                        if int(str(width).replace("px", "")) < 200:
                            continue
                    except ValueError:
                        pass
                if height:
                    try:
                        if int(str(height).replace("px", "")) < 100:
                            continue
                    except ValueError:
                        pass

                # 広告・アイコン系をスキップ
                src_lower = src.lower()
                skip_patterns = [
                    "logo", "icon", "avatar", "badge", "pixel",
                    "tracking", "ad-", "ads/", "advertisement",
                    "sponsor", "banner-ad", "1x1", "spacer",
                ]
                if any(pattern in src_lower for pattern in skip_patterns):
                    continue

                # alt/classでも広告判定
                alt = (img.get("alt") or "").lower()
                cls = (img.get("class") or [""])
                cls_str = " ".join(cls).lower() if isinstance(cls, list) else str(cls).lower()
                if any(p in alt for p in ["ad", "sponsor", "banner"]):
                    continue
                if any(p in cls_str for p in ["ad-", "ads", "sponsor"]):
                    continue

                return _make_absolute(src, article_url)

        return None

    except (requests.RequestException, Exception):
        return None


def _make_absolute(url: str, base_url: str) -> str:
    """相対URLを絶対URLに変換"""
    if url.startswith(("http://", "https://", "//")):
        if url.startswith("//"):
            return "https:" + url
        return url
    return urljoin(base_url, url)

#!/usr/bin/env python3
"""
Builds the static blog in docs/ from the .docx files in content/.

Run with:  python3 scripts/build.py

How posts work
---------------
- Every .docx file in content/ becomes one post.
- Optional filename prefix "YYYY-MM-DD-" sets the publish date and sort
  order, e.g. "2026-03-04-my-trip.docx". Without a prefix, the date is
  read from the Word file's own "created" metadata, and posts are still
  sorted newest-first.
- The post title is taken from (in order of preference): the docx's
  "Title" document property, the first Heading-1 paragraph in the file,
  or the filename.
- Images inside the .docx are extracted and placed in the same reading
  order they appear in the document.
- Arabic content is auto-detected and rendered right-to-left with an
  Arabic-friendly font; everything else renders left-to-right.
- Re-uploading a file with the exact same name updates that post in
  place. Renaming a file creates a new post with a new URL.
"""
import base64
import html
import json
import mimetypes
import os
import re
import shutil
import sys
from datetime import datetime, date
from pathlib import Path

try:
    import mammoth
except ImportError:
    sys.exit(
        "Missing dependency 'mammoth'. Install it with:\n"
        "    pip install mammoth python-docx"
    )

try:
    import docx as python_docx
except ImportError:
    python_docx = None  # metadata extraction is optional; we fall back gracefully

ROOT = Path(__file__).resolve().parent.parent
CONTENT_DIR = ROOT / "content"
DOCS_DIR = ROOT / "docs"
TEMPLATES_DIR = ROOT / "templates"
ASSETS_SRC = ROOT / "assets"
CONFIG_PATH = ROOT / "config.json"

ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]")
LETTER_RE = re.compile(r"[^\W\d_]", re.UNICODE)
DATE_PREFIX_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})-(.+)$")
TAG_RE = re.compile(r"<[^>]+>")
H1_RE = re.compile(r"<h1[^>]*>(.*?)</h1>", re.IGNORECASE | re.DOTALL)


def load_config():
    default = {
        "site_title": "My Blog",
        "site_tagline": "",
        "footer_text": "",
        "accent_color": "#b5643a",
        "posts_per_page": 12,
        "web3forms_access_key": "",
    }
    if CONFIG_PATH.exists():
        default.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    return default


def load_template(name):
    return (TEMPLATES_DIR / name).read_text(encoding="utf-8")


def fill(template, **kwargs):
    out = template
    for key, value in kwargs.items():
        out = out.replace("{{%s}}" % key, value)
    return out


def slugify_fallback(filename_stem, prefix_date):
    """Stable, ASCII-safe folder name for a post, even for Arabic filenames."""
    import hashlib
    digest = hashlib.md5(filename_stem.encode("utf-8")).hexdigest()[:8]
    date_part = prefix_date if prefix_date else "0000-00-00"
    return f"{date_part}-{digest}"


def humanize_filename(stem):
    text = stem.replace("_", " ").replace("-", " ").strip()
    return text[:1].upper() + text[1:] if text else "Untitled post"


def is_rtl_text(text):
    letters = LETTER_RE.findall(text)
    if not letters:
        return False
    arabic = ARABIC_RE.findall(text)
    return len(arabic) / max(len(letters), 1) > 0.35


def make_image_converter(images_dir, rel_prefix):
    images_dir.mkdir(parents=True, exist_ok=True)
    counter = {"n": 0}

    def convert_image(image):
        counter["n"] += 1
        ext = mimetypes.guess_extension(image.content_type or "") or ".png"
        if ext == ".jpe":
            ext = ".jpg"
        filename = f"img{counter['n']}{ext}"
        with image.open() as image_bytes:
            data = image_bytes.read()
        (images_dir / filename).write_bytes(data)
        return {"src": f"{rel_prefix}{filename}"}

    return mammoth.images.img_element(convert_image)


STYLE_MAP = """
p[style-name='Title'] => h1:fresh
p[style-name='Subtitle'] => p.subtitle:fresh
"""


def extract_docx_metadata(path):
    """Returns (title_or_None, created_date_or_None) from Word's document properties."""
    if python_docx is None:
        return None, None
    try:
        doc = python_docx.Document(str(path))
        props = doc.core_properties
        title = (props.title or "").strip() or None
        created = props.created or props.modified
        created_date = created.date() if isinstance(created, datetime) else created
        return title, created_date
    except Exception:
        return None, None


def strip_tags(html_text):
    return html.unescape(TAG_RE.sub(" ", html_text))


def make_excerpt(html_text, limit=180):
    text = re.sub(r"\s+", " ", strip_tags(html_text)).strip()
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0] + "\u2026"


def find_first_image(html_text):
    m = re.search(r'<img[^>]+src="([^"]+)"', html_text)
    return m.group(1) if m else None


def render_comment_section(config, title, rtl):
    key = (config.get("web3forms_access_key") or "").strip()
    labels = {
        "heading": "اترك رسالة" if rtl else "Leave a reply",
        "nickname": "اسمك (سيظهر مع رسالتك)" if rtl else "Your name",
        "message": "رسالتك" if rtl else "Your message",
        "send": "إرسال" if rtl else "Send",
        "sending": "جارٍ الإرسال…" if rtl else "Sending…",
        "success": "تم الإرسال، شكرًا لك!" if rtl else "Sent — thank you!",
        "error": "حدث خطأ، حاول مرة أخرى." if rtl else "Something went wrong — please try again.",
        "missing_name": "من فضلك أدخل اسمك." if rtl else "Please enter your name.",
    }

    if not key:
        note = (
            "التعليقات غير مفعّلة بعد." if rtl
            else "Comments aren't set up yet."
        )
        return (
            f'<section class="comment-section" dir="{"rtl" if rtl else "ltr"}">'
            f'<h2 class="comment-heading">{labels["heading"]}</h2>'
            f'<p class="comment-disabled">{note}</p>'
            f"</section>"
        )

    safe_title = title.replace('"', "&quot;")
    return f'''<section class="comment-section" dir="{"rtl" if rtl else "ltr"}">
      <h2 class="comment-heading">{labels["heading"]}</h2>
      <form class="comment-form" id="comment-form">
        <input type="hidden" name="access_key" value="{key}">
        <input type="hidden" name="subject" value="New reply on: {safe_title}">
        <input type="checkbox" name="botcheck" class="comment-botcheck" tabindex="-1" autocomplete="off">
        <div class="comment-field">
          <label for="comment-name">{labels["nickname"]}</label>
          <input type="text" id="comment-name" name="name" required maxlength="60">
        </div>
        <div class="comment-field">
          <label for="comment-message">{labels["message"]}</label>
          <textarea id="comment-message" name="message" rows="4" required maxlength="4000"></textarea>
        </div>
        <button type="submit">{labels["send"]}</button>
        <p class="comment-status" aria-live="polite"></p>
      </form>
      <script>
      (function () {{
        var form = document.getElementById('comment-form');
        var status = form.querySelector('.comment-status');
        form.addEventListener('submit', function (e) {{
          e.preventDefault();
          status.textContent = {labels["sending"]!r};
          status.className = 'comment-status';
          fetch('https://api.web3forms.com/submit', {{
            method: 'POST',
            headers: {{ 'Accept': 'application/json' }},
            body: new FormData(form)
          }}).then(function (r) {{ return r.json(); }}).then(function (data) {{
            if (data.success) {{
              status.textContent = {labels["success"]!r};
              status.className = 'comment-status comment-status-ok';
              form.reset();
            }} else {{
              status.textContent = {labels["error"]!r};
              status.className = 'comment-status comment-status-error';
            }}
          }}).catch(function () {{
            status.textContent = {labels["error"]!r};
            status.className = 'comment-status comment-status-error';
          }});
        }});
      }})();
      </script>
    </section>'''


def build_post(docx_path, config, css_version):
    stem = docx_path.stem
    date_match = DATE_PREFIX_RE.match(stem)
    prefix_date = date_match.group(1) if date_match else None
    name_part = date_match.group(2) if date_match else stem

    slug = slugify_fallback(name_part, prefix_date)
    post_dir = DOCS_DIR / "posts" / slug
    images_dir = post_dir / "images"
    if post_dir.exists():
        shutil.rmtree(post_dir)
    post_dir.mkdir(parents=True, exist_ok=True)

    meta_title, meta_date = extract_docx_metadata(docx_path)

    with open(docx_path, "rb") as f:
        result = mammoth.convert_to_html(
            f,
            style_map=STYLE_MAP,
            convert_image=make_image_converter(images_dir, "images/"),
        )
    content_html = result.value

    # Pull the title out of the first <h1> so it isn't shown twice.
    h1_match = H1_RE.search(content_html)
    if meta_title:
        title = meta_title
    elif h1_match:
        title = strip_tags(h1_match.group(1)).strip()
    else:
        title = humanize_filename(name_part)

    if h1_match:
        content_html = content_html[: h1_match.start()] + content_html[h1_match.end():]

    plain_text = strip_tags(content_html)
    rtl = is_rtl_text(title + " " + plain_text)
    direction = "rtl" if rtl else "ltr"
    lang = "ar" if rtl else "en"

    if prefix_date:
        post_date = prefix_date
    elif meta_date:
        post_date = meta_date.isoformat()
    else:
        post_date = date.fromtimestamp(docx_path.stat().st_mtime).isoformat()

    excerpt = make_excerpt(content_html)
    thumb_rel = find_first_image(content_html)
    comment_section = render_comment_section(config, title, rtl)

    post_html = fill(
        load_template("post.html"),
        TITLE=title,
        SITE_TITLE=config["site_title"],
        SITE_TAGLINE=config["site_tagline"],
        EXCERPT=excerpt,
        LANG=lang,
        DIR=direction,
        DATE=post_date,
        CONTENT=content_html,
        BACK_LABEL=("\u2192 عودة إلى كل المقالات" if rtl else "\u2190 Back to all posts"),
        FOOTER_TEXT=config["footer_text"],
        CSS_VERSION=css_version,
        COMMENT_SECTION=comment_section,
    )
    (post_dir / "index.html").write_text(post_html, encoding="utf-8")

    return {
        "slug": slug,
        "title": title,
        "date": post_date,
        "excerpt": excerpt,
        "dir": direction,
        "thumb": f"posts/{slug}/{thumb_rel}" if thumb_rel else None,
    }


def render_index(posts, config, css_version):
    posts_sorted = sorted(posts, key=lambda p: p["date"], reverse=True)
    if posts_sorted:
        cards = []
        for p in posts_sorted:
            thumb_html = (
                f'<div class="post-card-thumb"><img src="{p["thumb"]}" alt=""></div>'
                if p["thumb"]
                else ""
            )
            cards.append(
                fill(
                    load_template("post_card.html"),
                    DIR=p["dir"],
                    THUMB=thumb_html,
                    SLUG=p["slug"],
                    TITLE=p["title"],
                    DATE=p["date"],
                    EXCERPT=p["excerpt"],
                )
            )
        cards_html = "\n".join(cards)
    else:
        cards_html = '<li class="empty-state">No posts yet — upload a .docx file to content/ to publish your first one.</li>'

    index_html = fill(
        load_template("index.html"),
        SITE_TITLE=config["site_title"],
        SITE_TAGLINE=config["site_tagline"],
        POST_CARDS=cards_html,
        FOOTER_TEXT=config["footer_text"],
        CSS_VERSION=css_version,
    )
    (DOCS_DIR / "index.html").write_text(index_html, encoding="utf-8")


def copy_assets(config):
    import hashlib
    dest = DOCS_DIR / "assets"
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(ASSETS_SRC, dest)
    css_path = dest / "style.css"
    css = css_path.read_text(encoding="utf-8")
    css = css.replace("#b5643a", config["accent_color"])
    css_path.write_text(css, encoding="utf-8")
    # A short hash of the CSS, appended as ?v=... on the stylesheet link so
    # browsers fetch a fresh copy whenever the styling actually changes,
    # instead of reusing a stale cached version.
    return hashlib.md5(css.encode("utf-8")).hexdigest()[:8]


def main():
    config = load_config()
    DOCS_DIR.mkdir(exist_ok=True)

    css_version = copy_assets(config)

    posts_root = DOCS_DIR / "posts"
    if posts_root.exists():
        shutil.rmtree(posts_root)
    posts_root.mkdir(parents=True)

    docx_files = sorted(CONTENT_DIR.glob("*.docx"))
    posts = []
    for docx_path in docx_files:
        if docx_path.name.startswith("~$"):
            continue  # Word temp/lock file
        print(f"Building: {docx_path.name}")
        posts.append(build_post(docx_path, config, css_version))

    render_index(posts, config, css_version)
    (DOCS_DIR / ".nojekyll").touch()
    print(f"Done. Built {len(posts)} post(s) into docs/.")


if __name__ == "__main__":
    main()

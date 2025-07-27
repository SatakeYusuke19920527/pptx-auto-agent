# blueprints/create_pptx.py
import os
import json
import io
import traceback
import socket
from datetime import datetime
import zoneinfo
from pathlib import Path

from azure.functions import Blueprint, HttpRequest, HttpResponse  # type: ignore
from openai import AzureOpenAI
from openai import APITimeoutError, RateLimitError
import backoff

from pptx import Presentation
from pptx.util import Pt
from azure.storage.blob import BlobServiceClient, ContentSettings

# ─────────────────────────────
#  Blueprint
# ─────────────────────────────
create_pptx_bp = Blueprint()

# ─────────────────────────────
#  Azure OpenAI クライアント
# ─────────────────────────────
client = AzureOpenAI(
    api_version="2024-12-01-preview",
    azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
    api_key=os.environ["AZURE_OPENAI_KEY"],
    timeout=120,        # ソケット上限を 120 秒に延長
    max_retries=0       # SDK 内部リトライは無効化
)

SYSTEM_PROMPT = """
あなたはB2B向け資料作成のエキスパートです。
## 絶対ルール
- **出力は JSON だけ**。前後に説明やコードブロックを付けない。
- JSON のキーは "title" と "bullets" だけ。
- 箇条書きは必ず配列。
以下フォーマットで 5 枚:
[
  { "title": "タイトル1", "bullets": ["箇条書きA", "箇条書きB"] },
  ...
]
"""

# ─────────────────────────────
#  OpenAI 呼び出しを指数バックオフで包む
# ─────────────────────────────


@backoff.on_exception(
    backoff.expo,                               # 1→2→4→8→…
    (APITimeoutError, RateLimitError, socket.timeout),
    max_time=300                                # 最大 5 分で打ち切り
)
def fetch_outline(messages: list[dict]) -> list[dict]:
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        max_completion_tokens=400,
        timeout=180,
        temperature=0
    )
    return json.loads(resp.choices[0].message.content)

# ─────────────────────────────
#  /api/auto_ppt
# ─────────────────────────────


@create_pptx_bp.route(route="auto_ppt", methods=["GET"])
def auto_ppt(req: HttpRequest) -> HttpResponse:
    """
    1) Azure OpenAI でスライド構成生成
    2) template.pptx を読み込んで python-pptx で資料作成
    3) Blob Storage へアップロード
    4) 生成ファイルを返却
    """
    USER_PROMPT = (
        "テーマ: 日本人夫婦が行くドバイ旅行 "
        "対象読者: 旅行会社の営業担当者 "
        "目的: 人気観光地・アクティビティを 5 つ紹介し、各スライドを箇条書き 3 行にまとめる"
    )

    # 1. スライド構成生成（自動リトライ付き）
    slides = fetch_outline([
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user",   "content": USER_PROMPT},
    ])

    # 2. テンプレート読み込み & PPTX 構築
    template_path = Path(__file__).resolve().parent.parent / "template.pptx"
    prs = Presentation(str(template_path)
                       ) if template_path.exists() else Presentation()

    ts = datetime.now(zoneinfo.ZoneInfo("Asia/Tokyo")
                      ).strftime("%Y%m%d-%H%M%S")
    file_name = f"{ts}_auto_docs.pptx"

    # 表紙
    cover = prs.slides.add_slide(prs.slide_layouts[0])
    cover.shapes.title.text = slides[0]["title"]
    if cover.placeholders:
        cover.placeholders[1].text = f"Generated {ts}"

    # コンテンツ
    for slide in slides[1:]:
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = slide["title"]
        tf = s.shapes.placeholders[1].text_frame
        tf.clear()
        for b in slide["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(18)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    # 3. Blob アップロード
    try:
        blob_service = BlobServiceClient.from_connection_string(
            os.environ["BLOB_CONN"])
        container_name, blob_path = "pptstorage", f"generated/{file_name}"
        cc = blob_service.get_container_client(container_name)
        if not cc.exists():
            cc.create_container()

        cc.upload_blob(
            name=blob_path,
            data=buf.getvalue(),
            overwrite=True,
            content_settings=ContentSettings(
                content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
        )
        upload_status = f"https://{blob_service.account_name}.blob.core.windows.net/{container_name}/{blob_path}"
    except Exception as ex:
        upload_status = f"Blob upload failed: {ex}"

    # 4. HTTP 応答
    return HttpResponse(
        body=buf.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f'attachment; filename="{file_name}"',
            "X-Upload-Status": upload_status,
        },
    )

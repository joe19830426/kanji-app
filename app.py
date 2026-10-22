"""小学校 漢字小テスト自動作成アプリ（縦書き・A4/B4ヨコ・印刷対応）

起動:
    pip install streamlit anthropic python-docx
    streamlit run app.py

設定 (.streamlit/secrets.toml または環境変数):
    APP_PASSWORD      = "amida"   # 未設定なら認証なしで動作
    ANTHROPIC_API_KEY = "sk-ant-..."         # 画像/PDFの添付読み取りにだけ使用
"""

import base64
import hmac
import html
import io
import json
import math
import os
import random
import re

import streamlit as st
import streamlit.components.v1 as components

MODEL = "claude-sonnet-5-5"

PUBLISHERS = ["光村図書", "東京書籍", "教育出版", "学校図書", "三省堂", "その他"]
TERMS = ["1学期", "2学期", "3学期"]
PAPERS = {
    "A4ヨコ": (297, 210, "297mm 210mm"),
    "B4ヨコ": (364, 257, "364mm 257mm"),
}
KINDS = ["標準", "お助け（2画ヒント）", "ルビ付き", "お助け＋ルビ"]

# ---------------------------------------------------------------------------
# サンプルデータ  (答えの漢字語, よみ, 例文)
#   例文の {W} が答えの語の位置。[漢|かん] は「ルビ付き」のときだけルビになる。
#   各学年10語（学期・単元はコード側で自動割り当て）。実際の教科書に合わせて
#   画面の「追加登録」や添付プリントで増やしてください。
# ---------------------------------------------------------------------------
SAMPLE = {
    1: [
        ("学校", "がっこう", "ぼくの{W}は、[山|やま]の[下|した]だ。"),
        ("花火", "はなび", "みんなで{W}を[見|み]た。"),
        ("森林", "しんりん", "{W}の[中|なか]は、しずかだ。"),
        ("青空", "あおぞら", "{W}に[白|しろ]い雲がうかぶ。"),
        ("先生", "せんせい", "{W}に[本|ほん]をかりた。"),
        ("足音", "あしおと", "うしろから{W}がした。"),
        ("草花", "くさばな", "{W}に[水|みず]をやる。"),
        ("町", "まち", "[大|おお]きな{W}をあるく。"),
        ("天気", "てんき", "あしたも{W}がいい。"),
        ("左右", "さゆう", "{W}を[見|み]てわたる。"),
    ],
    2: [
        ("親切", "しんせつ", "{W}にして、よろこばれた。"),
        ("元気", "げんき", "あさから{W}いっぱいだ。"),
        ("海外", "かいがい", "{W}の[国|くに]に[行|い]く。"),
        ("通学", "つうがく", "{W}で[電車|でんしゃ]にのる。"),
        ("図書", "としょ", "{W}[室|しつ]で[本|ほん]を[読|よ]む。"),
        ("計算", "けいさん", "{W}のテストをする。"),
        ("家族", "かぞく", "{W}で[食|た]べる。"),
        ("黒板", "こくばん", "{W}に[字|じ]を[書|か]く。"),
        ("東西", "とうざい", "{W}に[長|なが]い[道|みち]がある。"),
        ("会話", "かいわ", "[友|とも]だちと{W}をする。"),
    ],
    3: [
        ("運動", "うんどう", "[毎朝|まいあさ]{W}をする。"),
        ("意見", "いけん", "[自分|じぶん]の{W}を[言|い]う。"),
        ("係", "かかり", "{W}の[仕事|しごと]をする。"),
        ("決心", "けっしん", "{W}して、[勉強|べんきょう]する。"),
        ("湖", "みずうみ", "{W}の[水|みず]がすんでいる。"),
        ("研究", "けんきゅう", "{W}を[進|すす]める。"),
        ("階段", "かいだん", "{W}を[上|のぼ]る。"),
        ("整理", "せいり", "[本|ほん]だなを{W}する。"),
        ("写真", "しゃしん", "[家族|かぞく]の{W}をとる。"),
        ("勉強", "べんきょう", "[算数|さんすう]を{W}する。"),
    ],
    4: [
        ("博物館", "はくぶつかん", "[町|まち]の{W}に[行|い]く。"),
        ("最近", "さいきん", "{W}、[本|ほん]をよく[読|よ]む。"),
        ("共同", "きょうどう", "みんなで{W}して[作|つく]る。"),
        ("努力", "どりょく", "{W}をつづける。"),
        ("信号", "しんごう", "{W}が[赤|あか]にかわる。"),
        ("給食", "きゅうしょく", "{W}のじかんだ。"),
        ("観察", "かんさつ", "[花|はな]を{W}する。"),
        ("改札", "かいさつ", "[駅|えき]の{W}を[通|とお]る。"),
        ("希望", "きぼう", "[大|おお]きな{W}をもつ。"),
        ("季節", "きせつ", "[春|はる]は、いい{W}だ。"),
    ],
    5: [
        ("経験", "けいけん", "[自分|じぶん]の{W}を[話|はな]す。"),
        ("資料", "しりょう", "{W}を[集|あつ]める。"),
        ("規則", "きそく", "{W}を[守|まも]る。"),
        ("可能", "かのう", "{W}かどうか[考|かんが]える。"),
        ("確認", "かくにん", "[答|こた]えを{W}する。"),
        ("現在", "げんざい", "{W}の[時間|じかん]を[知|し]る。"),
        ("情報", "じょうほう", "{W}を[調|しら]べる。"),
        ("移動", "いどう", "[遠|とお]くへ{W}する。"),
        ("構成", "こうせい", "{W}を[考|かんが]える。"),
        ("提案", "ていあん", "[新|あたら]しい{W}をする。"),
    ],
    6: [
        ("宇宙", "うちゅう", "{W}を[旅|たび]する。"),
        ("貴重", "きちょう", "{W}な[資料|しりょう]を[見|み]る。"),
        ("郵便", "ゆうびん", "{W}で[手紙|てがみ]を[送|おく]る。"),
        ("認識", "にんしき", "[正|ただ]しく{W}する。"),
        ("誠実", "せいじつ", "{W}な[人|ひと]だ。"),
        ("縦横", "じゅうおう", "{W}に[線|せん]を[引|ひ]く。"),
        ("宝物", "たからもの", "{W}を[見|み]つけた。"),
        ("展覧会", "てんらんかい", "{W}に[行|い]く。"),
        ("憲法", "けんぽう", "{W}を[学|まな]ぶ。"),
        ("磁石", "じしゃく", "{W}で[鉄|てつ]をくっつける。"),
    ],
}

# 語の1字目の画数（お助けヒント用）
STROKES = {
    "学": 8, "花": 7, "森": 12, "青": 8, "先": 6, "足": 7, "草": 9, "町": 7, "天": 4, "左": 5,
    "親": 16, "元": 4, "海": 9, "通": 10, "図": 7, "計": 9, "家": 10, "黒": 11, "東": 8, "会": 6,
    "運": 12, "意": 13, "係": 9, "決": 7, "湖": 12, "研": 9, "階": 12, "整": 16, "写": 5, "勉": 10,
    "博": 12, "最": 12, "共": 6, "努": 7, "信": 9, "給": 12, "観": 18, "改": 7, "希": 7, "季": 8,
    "経": 11, "資": 13, "規": 11, "可": 5, "確": 15, "現": 11, "情": 11, "移": 11, "構": 14, "提": 12,
    "宇": 6, "貴": 12, "郵": 11, "認": 14, "誠": 13, "縦": 16, "宝": 8, "展": 10, "憲": 16, "磁": 14,
}


def build_samples():
    out = []
    for grade, rows in SAMPLE.items():
        for i, (word, reading, sent) in enumerate(rows):
            out.append({
                "grade": grade,
                "term": TERMS[(i * 3) // len(rows)],
                "unit": f"単元{i // 2 + 1}",
                "word": word, "reading": reading, "sentence": sent,
            })
    return out


SAMPLE_ENTRIES = build_samples()

# ---------------------------------------------------------------------------
# 認証（APP_PASSWORD を設定した場合のみ）
# ---------------------------------------------------------------------------


def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    expected = st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", ""))
    if not expected:
        return True
    st.title("🔒 漢字小テスト作成")
    with st.form("login"):
        pw = st.text_input("パスワード", type="password")
        ok = st.form_submit_button("ログイン")
    if ok:
        if hmac.compare_digest(pw.encode(), str(expected).encode()):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが違います。")
    return False


# ---------------------------------------------------------------------------
# 添付プリント／追加登録の読み取り
# ---------------------------------------------------------------------------
MARK_RE = re.compile(r"【([^|｜】]+)(?:[|｜]([^】]+))?】")
NUM_RE = re.compile(r"^\s*[\(（]?[0-9０-９①-⑳]+[\)）.．、]?\s*")


def parse_text(text: str, unit="添付プリント"):
    """行ごとに【語|よみ】を含む文、または「語,よみ,例文」形式を読み取る。"""
    entries = []
    for raw in text.splitlines():
        line = NUM_RE.sub("", raw.strip())
        if not line:
            continue
        m = MARK_RE.search(line)
        if m:
            entries.append({
                "term": "添付", "unit": unit, "word": m.group(1).strip(),
                "reading": (m.group(2) or "").strip(),
                "sentence": line[:m.start()] + "{W}" + line[m.end():],
            })
            continue
        parts = [p.strip() for p in re.split(r"[,，\t]", line)]
        if len(parts) >= 3 and parts[0] and parts[0] in parts[2]:
            entries.append({
                "term": "添付", "unit": unit, "word": parts[0], "reading": parts[1],
                "sentence": parts[2].replace(parts[0], "{W}", 1),
            })
    return entries


def decode_bytes(b: bytes) -> str:
    for enc in ("utf-8-sig", "cp932", "utf-16"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="ignore")


def read_with_claude(data: bytes, media_type: str):
    import anthropic
    key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", "")).strip()
    if not key:
        raise RuntimeError("画像・PDFの読み取りには ANTHROPIC_API_KEY の設定が必要です。")
    b64 = base64.standard_b64encode(data).decode()
    block = ({"type": "document", "source": {"type": "base64", "media_type": media_type, "data": b64}}
             if media_type == "application/pdf" else
             {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}})
    prompt = (
        "添付は小学校の漢字テストのプリントです。書き問題（ひらがなを漢字に直す問題）を全て読み取り、"
        "読み問題の場合は下線部の語を答えの語として同様に扱ってください。"
        "次のJSON配列だけを出力してください（説明やコードブロックは不要）。\n"
        '[{"sentence":"文全体。答えの漢字語の部分だけを {W} に置き換える","word":"答えの漢字語","reading":"ひらがなのよみ"}]'
    )
    client = anthropic.Anthropic(api_key=key)
    resp = client.messages.create(
        model=MODEL, max_tokens=8000,
        messages=[{"role": "user", "content": [block, {"type": "text", "text": prompt}]}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    m = re.search(r"\[.*\]", text, re.DOTALL)
    if not m:
        raise ValueError("プリントから問題を読み取れませんでした。")
    rows = json.loads(m.group(0))
    return [{"term": "添付", "unit": "添付プリント", "word": r["word"],
             "reading": r.get("reading", ""), "sentence": r["sentence"]}
            for r in rows if r.get("word") and "{W}" in r.get("sentence", "")]


def read_attachment(f):
    name = f.name.lower()
    data = f.getvalue()
    if name.endswith((".txt", ".csv", ".md")):
        return parse_text(decode_bytes(data))
    if name.endswith(".docx"):
        import docx  # python-docx
        d = docx.Document(io.BytesIO(data))
        lines = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                lines.append(",".join(c.text for c in row.cells))
        return parse_text("\n".join(lines))
    if name.endswith(".pdf"):
        return read_with_claude(data, "application/pdf")
    if name.endswith((".png", ".jpg", ".jpeg")):
        return read_with_claude(data, "image/png" if name.endswith(".png") else "image/jpeg")
    raise ValueError("対応していない形式です。")


# ---------------------------------------------------------------------------
# プリントHTML生成
# ---------------------------------------------------------------------------
CSS = """
*{box-sizing:border-box}
html,body{margin:0}
body{background:#e9ecef;color:#111;
 font-family:"UD デジタル 教科書体 NK-R","UD Digi Kyokasho NK-R","Hiragino Mincho ProN","Yu Mincho","BIZ UDMincho",serif}
.toolbar{position:sticky;top:0;z-index:5;display:flex;gap:8px;align-items:center;padding:8px;
 background:#fff;border-bottom:1px solid #ccc;font-family:sans-serif;font-size:13px;color:#555}
.toolbar button{font-size:16px;padding:10px 18px;min-height:44px;border:0;border-radius:8px;background:#1a73e8;color:#fff}
.scaler{transform-origin:top left;padding:8px}
.sheet{width:__W__mm;height:__H__mm;background:#fff;padding:__M__mm;margin:0 0 8px;display:flex;
 flex-direction:column;box-shadow:0 1px 6px #0003;overflow:hidden}
.head{display:flex;justify-content:space-between;align-items:flex-start;gap:6mm;
 border-bottom:.5mm solid #111;padding-bottom:2mm;margin-bottom:3mm}
.head h1{font-size:6.4mm;margin:0 0 1.2mm;letter-spacing:.3mm}
.head p{font-size:3.5mm;margin:0}
.score{flex:none;width:34mm;height:14mm;border:.4mm solid #111;border-radius:2mm;font-size:3.6mm;
 display:flex;align-items:flex-end;justify-content:flex-end;padding:1mm 2mm}
.body{flex:1;min-height:0;display:grid;direction:rtl;grid-template-columns:auto 1fr;gap:4mm}
.names{direction:ltr;writing-mode:vertical-rl;height:100%}
.nf{display:flex;inline-size:100%;border:.35mm solid #111;margin-left:2.5mm;font-size:4.2mm}
.nf.name{block-size:13mm}.nf.cls,.nf.num{block-size:9mm}
.nf .lb{padding:1.2mm 0;text-align:center;letter-spacing:1mm}
.nf .fill{flex:1;border-top:.35mm solid #111}
.qs{direction:rtl;display:grid;grid-auto-flow:column;min-height:0;min-width:0}
.q{direction:ltr;writing-mode:vertical-rl;display:flex;gap:1.5mm;border-left:.2mm dotted #999;
 padding:1mm .8mm;font-size:var(--fs);line-height:2.1;min-height:0}
.q .sent{flex:none;white-space:nowrap}
.no{display:inline-block;flex:none;width:5.2mm;height:5.2mm;border:.3mm solid #111;border-radius:50%;
 font-size:3.2mm;line-height:5mm;text-align:center;font-family:sans-serif}
.tcy{text-combine-upright:all;-webkit-text-combine:horizontal}
rt{font-size:.48em;line-height:1}
.bx{display:inline-block;width:var(--bx);height:var(--bx);border:.35mm solid #111;margin:.5mm 0;
 line-height:calc(var(--bx) - .8mm);text-align:center;font-size:calc(var(--bx)*.78);color:#c4001a;vertical-align:top}
.hint{margin-inline-start:auto;flex:none;display:flex;gap:.5mm;align-items:center;
 border:.3mm dashed #666;border-radius:1mm;padding:.6mm;font-size:3mm;color:#555;line-height:1.2}
.hk{display:inline-block;width:6.4mm;height:6.4mm;line-height:6.4mm;font-size:5.4mm;text-align:center;
 color:#555;-webkit-mask-image:linear-gradient(135deg,#000 0%,#000 42%,transparent 58%);
 mask-image:linear-gradient(135deg,#000 0%,#000 42%,transparent 58%)}
@page{size:__PAPER__ landscape;margin:0}
@media print{
 body{background:#fff}
 .toolbar{display:none}
 .scaler{transform:none!important;padding:0;width:auto!important}
 .sheet{height:calc(__H__mm - 1mm);margin:0;box-shadow:none;break-after:page;page-break-after:always}
 .sheet:last-child{break-after:auto;page-break-after:auto}
 *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
"""

JS = """
function fit(){var s=document.querySelector('.sheet'),c=document.querySelector('.scaler');
 c.style.transform='none';var w=s.offsetWidth+16,k=Math.min(1,window.innerWidth/w);
 c.style.transform='scale('+k+')';}
window.addEventListener('load',fit);window.addEventListener('resize',fit);
window.onbeforeprint=function(){document.querySelector('.scaler').style.transform='none'};
window.onafterprint=fit;
"""

RUBY_RE = re.compile(r"\[([^\]|]+)\|([^\]]+)\]")


def sentence_html(e, ruby: bool, answer: bool) -> str:
    s = html.escape(e["sentence"])
    s = RUBY_RE.sub(
        lambda m: f"<ruby>{m.group(1)}<rt>{m.group(2)}</rt></ruby>" if ruby else m.group(1), s)
    boxes = "".join(f'<span class="bx">{html.escape(c) if answer else ""}</span>' for c in e["word"])
    rt = f"<rt>{html.escape(e['reading'])}</rt>" if e.get("reading") else ""
    target = f'<ruby class="t">{boxes}{rt}</ruby>'
    return s.replace("{W}", target) if "{W}" in s else s + target


def cell_html(i, e, help_on, ruby, answer) -> str:
    no = f'<span class="no"><span class="tcy">{i}</span></span>'
    hint = ""
    if help_on and not answer:
        first = e["word"][0]
        st_n = STROKES.get(first)
        cnt = f'<span class="hs"><span class="tcy">{st_n}</span>画</span>' if st_n else ""
        hint = f'<span class="hint"><span class="hk">{html.escape(first)}</span>{cnt}</span>'
    return f'<div class="q">{no}<span class="sent">{sentence_html(e, ruby, answer)}</span>{hint}</div>'


def sheet_html(items, title, kind, paper, answer) -> str:
    help_on = kind in ("お助け（2画ヒント）", "お助け＋ルビ")
    ruby = kind in ("ルビ付き", "お助け＋ルビ")
    n = len(items)
    rows = 1 if n <= 25 else 2
    cols = math.ceil(n / rows)
    fs, bx = ("4.6mm", "7.2mm") if rows == 1 else ("3.9mm", "6.2mm")
    pts = f"{100 / n:.1f}".rstrip("0").rstrip(".")
    cells = "".join(cell_html(i + 1, e, help_on, ruby, answer) for i, e in enumerate(items))
    ttl = html.escape(title) + ("（解答）" if answer else "（100点満点）")
    desc = ("赤字が答えです。" if answer else
            f"文中の　線のひらがなを漢字に直して、□に1字ずつ書きましょう。（1問{pts}点）")
    names = ('<div class="names"><div class="nf name"><span class="lb">名前</span><span class="fill"></span></div>'
             '<div class="nf cls"><span class="lb">組</span><span class="fill"></span></div>'
             '<div class="nf num"><span class="lb">番</span><span class="fill"></span></div></div>')
    return (f'<section class="sheet"><div class="head"><div><h1>{ttl}</h1><p>{desc}</p></div>'
            f'<div class="score">点</div></div><div class="body">{names}'
            f'<div class="qs" style="--fs:{fs};--bx:{bx};grid-template-columns:repeat({cols},1fr);'
            f'grid-template-rows:repeat({rows},1fr)">{cells}</div></div></section>')


def document_html(sheets, paper) -> str:
    w, h, size = PAPERS[paper]
    css = (CSS.replace("__W__", str(w)).replace("__H__", str(h))
           .replace("__M__", "8").replace("__PAPER__", size))
    return (f'<!doctype html><html lang="ja"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1"><style>{css}</style></head><body>'
            f'<div class="toolbar"><button onclick="window.print()">🖨 共有・プリント</button>'
            f'<span>{paper}・横向きで印刷（倍率100%／余白なし）</span></div>'
            f'<div class="scaler">{"".join(sheets)}</div><script>{JS}</script></body></html>')


# ---------------------------------------------------------------------------
# 問題の抽出
# ---------------------------------------------------------------------------
def pick(pool, n, seed):
    rng = random.Random(seed)
    out = []
    while pool and len(out) < n:
        cyc = pool[:]
        rng.shuffle(cyc)
        out += cyc[: n - len(out)]
    return out


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="漢字小テスト作成", page_icon="📝", layout="wide")
    if not check_password():
        return
    st.markdown("<style>.stButton>button,.stDownloadButton>button{min-height:46px}"
                "div[role=radiogroup]{gap:.4rem}</style>", unsafe_allow_html=True)
    st.title("📝 漢字小テスト作成（縦書き）")

    if "seed" not in st.session_state:
        st.session_state["seed"] = random.randint(0, 10**6)

    with st.expander("⚙ 設定", expanded=True):
        c1, c2, c3 = st.columns(3)
        grade = c1.selectbox("学年", [1, 2, 3, 4, 5, 6], format_func=lambda g: f"{g}年生")
        publisher = c2.selectbox("教科書会社", PUBLISHERS)
        paper = c3.selectbox("用紙", list(PAPERS))

        c4, c5 = st.columns(2)
        mode = c4.radio("漢字の選び方", ["学期の範囲からランダムに選ぶ", "全登録文字から抽出する"])
        n = c5.slider("問題数", 10, 50, 20)

        term, units = None, []
        if mode.startswith("学期"):
            c6, c7 = st.columns(2)
            term = c6.selectbox("学期", TERMS)
            avail = sorted({e["unit"] for e in SAMPLE_ENTRIES
                            if e["grade"] == grade and e["term"] == term})
            units = c7.multiselect("単元（未選択＝学期全体）", avail)

        title = st.text_input("プリント名", value=f"{grade}年 漢字小テスト")
        kind = st.radio("プリントの種類", KINDS, horizontal=True)

        extra_text = st.text_area(
            "追加登録（任意）", height=90,
            placeholder="1行に1問。答えの語を【】で囲む。よみは【語|よみ】\n例：夏休みに【海水浴|かいすいよく】へ行く。")
        up = st.file_uploader("既存プリントを添付（txt / csv / docx / 画像 / PDF）",
                              type=["txt", "csv", "md", "docx", "png", "jpg", "jpeg", "pdf"])
        use_mode = st.radio("添付・追加登録の使い方",
                            ["サンプルに加えて抽選する", "添付・追加登録の問題だけをそのまま使う"],
                            horizontal=True)

        if st.button("🔄 問題を再抽選"):
            st.session_state["seed"] = random.randint(0, 10**6)

    # --- 問題の準備 ---
    extra = parse_text(extra_text, unit="追加登録") if extra_text.strip() else []
    if up is not None:
        try:
            att = read_attachment(up)
            if att:
                st.success(f"添付プリントから {len(att)} 問を読み取りました。")
            else:
                st.warning("添付プリントから問題を読み取れませんでした。"
                           "【語|よみ】形式の txt/csv か、画像・PDFをお試しください。")
            extra += att
        except Exception as ex:
            st.error(f"添付ファイルの読み取りに失敗しました: {ex}")

    base = [e for e in SAMPLE_ENTRIES if e["grade"] == grade]
    if term:
        base = [e for e in base if e["term"] == term and (not units or e["unit"] in units)]

    if use_mode.startswith("添付") and extra:
        items = extra[:n]
        dup = False
    else:
        pool = base + extra
        items = pick(pool, n, st.session_state["seed"])
        dup = len(pool) < n
    if not items:
        st.warning("出題できる問題がありません。条件を見直してください。")
        return
    if dup:
        st.info(f"登録数（{len(pool)}問）が問題数より少ないため、一部の問題が重複しています。"
                "追加登録や添付で増やせます。")

    q_html = sheet_html(items, title, kind, paper, answer=False)
    a_html = sheet_html(items, title, kind, paper, answer=True)

    w, h, _ = PAPERS[paper]
    page_px = int(h * 3.78) + 24

    tab_q, tab_a, tab_b = st.tabs(["問題のみ", "解答のみ", "問題＋解答"])
    with tab_q:
        components.html(document_html([q_html], paper), height=page_px + 70, scrolling=True)
    with tab_a:
        components.html(document_html([a_html], paper), height=page_px + 70, scrolling=True)
    with tab_b:
        components.html(document_html([q_html, a_html], paper), height=page_px * 2 + 70, scrolling=True)

    st.download_button("📥 HTMLを保存（ブラウザで開いて印刷）",
                       document_html([q_html, a_html], paper),
                       file_name="kanji_test.html", mime="text/html")
    st.caption(f"{publisher}／{grade}年　※サンプル例文は学年配当漢字に合わせた練習用データです。"
               "実際の教科書の単元に合わせるには、追加登録か添付をご利用ください。")


main()

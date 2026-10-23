"""小学校 漢字小テスト自動作成アプリ（縦書き・A4/B4ヨコ・印刷対応）

起動:
    pip install streamlit anthropic python-docx
    streamlit run app.py

設定 (.streamlit/secrets.toml または環境変数):
    APP_PASSWORD      = "amida"   # 未設定なら認証なしで動作
    ANTHROPIC_API_KEY = "sk-ant-..."         # 問題文の自動作成・画像/PDFの読み取りに使用
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

MODEL = "claude-3-7-sonnet-20250219"

PUBLISHERS = ["光村図書", "東京書籍", "教育出版", "学校図書", "三省堂", "その他"]
TERMS = ["1学期", "2学期", "3学期"]
COUNTS = [10, 20, 30, 40, 50]
PAPERS = {
    "A4ヨコ": (297, 210, "297mm 210mm"),
    "B4ヨコ": (364, 257, "364mm 257mm"),
}
KINDS = ["標準", "お助け（2画ヒント）", "ルビ付き", "お助け＋ルビ"]

# ---------------------------------------------------------------------------
# 学年別漢字配当表（小学校学習指導要領 平成29年告示・全1,026字）
# 1年80字 / 2年160字 / 3年200字 / 4年202字 / 5年193字 / 6年191字
# ---------------------------------------------------------------------------
GRADE_KANJI = {
    1: "一右雨円王音下火花貝学気九休玉金空月犬見五口校左三山子四糸字耳七車手十出女小上森人水正生青夕石赤千川先早草足村大男竹中虫町天田土二日入年白八百文木本名目立力林六",
    2: "引羽雲園遠何科夏家歌画回会海絵外角楽活間丸岩顔汽記帰弓牛魚京強教近兄形計元言原戸古午後語工公広交光考行高黄合谷国黒今才細作算止市矢姉思紙寺自時室社弱首秋週春書少場色食心新親図数西声星晴切雪船線前組走多太体台地池知茶昼長鳥朝直通弟店点電刀冬当東答頭同道読内南肉馬売買麦半番父風分聞米歩母方北毎妹万明鳴毛門夜野友用曜来里理話",
    3: "悪安暗医委意育員院飲運泳駅央横屋温化荷界開階寒感漢館岸起期客究急級宮球去橋業曲局銀区苦具君係軽血決研県庫湖向幸港号根祭皿仕死使始指歯詩次事持式実写者主守取酒受州拾終習集住重宿所暑助昭消商章勝乗植申身神真深進世整昔全相送想息速族他打対待代第題炭短談着注柱丁帳調追定庭笛鉄転都度投豆島湯登等動童農波配倍箱畑発反坂板皮悲美鼻筆氷表秒病品負部服福物平返勉放味命面問役薬由油有遊予羊洋葉陽様落流旅両緑礼列練路和",
    4: "愛案以衣位茨印英栄媛塩岡億加果貨課芽賀改械害街各覚潟完官管関観願岐希季旗器機議求泣給挙漁共協鏡競極熊訓軍郡群径景芸欠結建健験固功好香候康佐差菜最埼材崎昨札刷察参産散残氏司試児治滋辞鹿失借種周祝順初松笑唱焼照城縄臣信井成省清静席積折節説浅戦選然争倉巣束側続卒孫帯隊達単置仲沖兆低底的典伝徒努灯働特徳栃奈梨熱念敗梅博阪飯飛必票標不夫付府阜富副兵別辺変便包法望牧末満未民無約勇要養浴利陸良料量輪類令冷例連老労録",
    5: "圧囲移因永営衛易益液演応往桜可仮価河過快解格確額刊幹慣眼紀基寄規喜技義逆久旧救居許境均禁句型経潔件険検限現減故個護効厚耕航鉱構興講告混査再災妻採際在財罪殺雑酸賛士支史志枝師資飼示似識質舎謝授修述術準序招証象賞条状常情織職制性政勢精製税責績接設絶祖素総造像増則測属率損貸態団断築貯張停提程適統堂銅導得毒独任燃能破犯判版比肥非費備評貧布婦武復複仏粉編弁保墓報豊防貿暴脈務夢迷綿輸余容略留領歴",
    6: "胃異遺域宇映延沿恩我灰拡革閣割株干巻看簡危机揮貴疑吸供胸郷勤筋系敬警劇激穴券絹権憲源厳己呼誤后孝皇紅降鋼刻穀骨困砂座済裁策冊蚕至私姿視詞誌磁射捨尺若樹収宗就衆従縦縮熟純処署諸除承将傷障蒸針仁垂推寸盛聖誠舌宣専泉洗染銭善奏窓創装層操蔵臓存尊退宅担探誕段暖値宙忠著庁頂腸潮賃痛敵展討党糖届難乳認納脳派拝背肺俳班晩否批秘俵腹奮並陛閉片補暮宝訪亡忘棒枚幕密盟模訳郵優預幼欲翌乱卵覧裏律臨朗論",
}

# ---------------------------------------------------------------------------
# サンプル問題（APIキー未設定時の代替・補充用） (答えの語, よみ, 例文)
#   例文の {W} が答えの位置。[漢|かん] は「ルビ付き」のときだけルビになる。
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
        ("見える", "みえる", "とおくの[山|やま]が{W}。"),
        ("休む", "やすむ", "きょうは[学校|がっこう]を{W}。"),
    ],
    2: [
        ("親切", "しんせつ", "{W}にして、よろこばれた。"),
        ("元気", "げんき", "あさから{W}いっぱいだ。"),
        ("海外", "かいがい", "{W}の[国|くに]に[行|い]く。"),
        ("通学", "つうがく", "{W}で[電車|でんしゃ]にのる。"),
        ("図書", "としょ", "{W}[室|しつ]で[本|ほん]を[読|よ]む。"),
        ("計算", "けいさん", "{W}のテストをする。"),
        ("家", "いえ", "[大|おお]きな{W}にすむ。"),
        ("黒", "くろ", "{W}い[犬|いぬ]がいる。"),
        ("東西", "とうざい", "{W}に[長|なが]い[道|みち]がある。"),
        ("会話", "かいわ", "[友|とも]だちと{W}をする。"),
        ("答える", "こたえる", "[先生|せんせい]に{W}。"),
        ("遠い", "とおい", "[学校|がっこう]は{W}。"),
    ],
    3: [
        ("運動", "うんどう", "[毎朝|まいあさ]{W}をする。"),
        ("意見", "いけん", "[自分|じぶん]の{W}を[言|い]う。"),
        ("係", "かかり", "{W}の[仕事|しごと]をする。"),
        ("決心", "けっしん", "{W}して、[勉強|べんきょう]する。"),
        ("湖", "みずうみ", "{W}の[水|みず]がすんでいる。"),
        ("研究", "けんきゅう", "{W}を[進|すす]める。"),
        ("階", "かい", "[二|に]{W}の[教室|きょうしつ]へ[行|い]く。"),
        ("整理", "せいり", "[本|ほん]だなを{W}する。"),
        ("写真", "しゃしん", "[家族|かぞく]の{W}をとる。"),
        ("勉強", "べんきょう", "[算数|さんすう]を{W}する。"),
        ("決める", "きめる", "[係|かかり]を{W}。"),
        ("住む", "すむ", "[町|まち]に{W}。"),
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
        ("浅い", "あさい", "[川|かわ]が{W}。"),
        ("働く", "はたらく", "[父|ちち]は[会社|かいしゃ]で{W}。"),
    ],
    5: [
        ("経験", "けいけん", "[自分|じぶん]の{W}を[話|はな]す。"),
        ("資料", "しりょう", "{W}を[集|あつ]める。"),
        ("規則", "きそく", "{W}を[守|まも]る。"),
        ("可能", "かのう", "{W}かどうか[考|かんが]える。"),
        ("在庫", "ざいこ", "[店|みせ]の{W}を[調|しら]べる。"),
        ("現在", "げんざい", "{W}の[時間|じかん]を[知|し]る。"),
        ("情報", "じょうほう", "{W}を[調|しら]べる。"),
        ("移動", "いどう", "[遠|とお]くへ{W}する。"),
        ("構成", "こうせい", "{W}を[考|かんが]える。"),
        ("提案", "ていあん", "[新|あたら]しい{W}をする。"),
        ("任せる", "まかせる", "[仕事|しごと]を{W}。"),
        ("確かめる", "たしかめる", "[答|こた]えを{W}。"),
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
        ("訪ねる", "たずねる", "[友|とも]だちを{W}。"),
        ("忘れる", "わすれる", "[宿題|しゅくだい]を{W}。"),
    ],
}

KANJI_RE = re.compile(r"[\u4e00-\u9fff々]")


def sample_pool(grade):
    """その学年までの内蔵サンプル問題"""
    return [{"word": w, "reading": r, "sentence": s}
            for g in range(1, grade + 1) for w, r, s in SAMPLE[g]]


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
# Claude API
# ---------------------------------------------------------------------------
def api_key() -> str:
    return str(st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", ""))).strip()


def call_claude(content, max_tokens=8000) -> str:
    import anthropic
    key = api_key()
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY が設定されていません。")
    client = anthropic.Anthropic(api_key=key)
    resp = client.messages.create(model=MODEL, max_tokens=max_tokens,
                                  messages=[{"role": "user", "content": content}])
    return "".join(b.text for b in resp.content if b.type == "text")


def parse_json_list(text: str):
    m = re.search(r"\[.*\]", text, re.DOTALL)
    if not m:
        raise ValueError("AIの応答からJSONを取得できませんでした。")
    return json.loads(m.group(0))


def generate_with_claude(grade, kanji, n):
    prompt = f"""小学{grade}年生向けの漢字書き取りテストの問題を、ちょうど{n}問作ってください。
出題に使う漢字（この中から選ぶ）: {kanji}

【ルール】
1. 各問は「答えとなる語（漢字で書く）」を1つ決める。答えの語には、上の出題漢字を少なくとも1字含める。
2. 答えの語に使う漢字・文中の漢字は、小学{grade}年生までに習う漢字だけにする。習っていない漢字はひらがなにする。
3. 熟語（2字以上）、1字の語、送り仮名のある語（例：任せる・細かい）をバランスよく混ぜる。送り仮名のある語は2割ほど。
4. 例文は{grade}年生が読める短い文（答えを除いて15字以内）。答えの部分を {{W}} と書く。
5. 文中の答え以外の漢字は [漢字|よみ] の形でひらがなの読みを付ける（例：[先生|せんせい]に{{W}}。）。
6. 出題漢字がなるべく偏らないようにし、同じ語を重複させない。易しい問題から順に並べる。
JSON配列だけを出力（説明やコードブロックは不要）:
[{{"kanji":"出題漢字1字","word":"答えの語","reading":"ひらがな","sentence":"例文"}}]"""
    rows = parse_json_list(call_claude(prompt))
    allowed = set("".join(GRADE_KANJI[g] for g in range(1, grade + 1)))
    out = []
    for r in rows:
        w, s = str(r.get("word", "")).strip(), str(r.get("sentence", ""))
        if not w or "{W}" not in s:
            continue
        if any(KANJI_RE.match(c) and c not in allowed for c in w):
            continue
        out.append({"word": w, "reading": str(r.get("reading", "")).strip(), "sentence": s})
    return out


# ---------------------------------------------------------------------------
# 添付プリントの読み取り
# ---------------------------------------------------------------------------
MARK_RE = re.compile(r"【([^|｜】]+)(?:[|｜]([^】]+))?】")
NUM_RE = re.compile(r"^\s*[\(（]?[0-9０-９①-⑳]+[\)）.．、]?\s*")


def parse_text(text: str):
    """行ごとに【語|よみ】を含む文、または「語,よみ,例文」形式を読み取る。"""
    entries = []
    for raw in text.splitlines():
        line = NUM_RE.sub("", raw.strip())
        if not line:
            continue
        m = MARK_RE.search(line)
        if m:
            entries.append({"word": m.group(1).strip(), "reading": (m.group(2) or "").strip(),
                            "sentence": line[:m.start()] + "{W}" + line[m.end():]})
            continue
        parts = [p.strip() for p in re.split(r"[,，\t]", line)]
        if len(parts) >= 3 and parts[0] and parts[0] in parts[2]:
            entries.append({"word": parts[0], "reading": parts[1],
                            "sentence": parts[2].replace(parts[0], "{W}", 1)})
    return entries


def decode_bytes(b: bytes) -> str:
    for enc in ("utf-8-sig", "cp932", "utf-16"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", errors="ignore")


def read_with_claude(data: bytes, media_type: str):
    b64 = base64.standard_b64encode(data).decode()
    kind = "document" if media_type == "application/pdf" else "image"
    block = {"type": kind, "source": {"type": "base64", "media_type": media_type, "data": b64}}
    prompt = ("添付は小学校の漢字テストのプリントです。書き問題（ひらがなを漢字に直す問題）を全て読み取り、"
              "読み問題の場合は下線部の語を答えの語として同様に扱ってください。"
              "次のJSON配列だけを出力してください（説明やコードブロックは不要）。\n"
              '[{"sentence":"文全体。答えの漢字語の部分だけを {W} に置き換える","word":"答えの漢字語","reading":"ひらがなのよみ"}]')
    rows = parse_json_list(call_claude([block, {"type": "text", "text": prompt}]))
    return [{"word": r["word"], "reading": r.get("reading", ""), "sentence": r["sentence"]}
            for r in rows if r.get("word") and "{W}" in r.get("sentence", "")]


def read_attachment(f):
    name, data = f.name.lower(), f.getvalue()
    if name.endswith((".txt", ".csv", ".md")):
        return parse_text(decode_bytes(data))
    if name.endswith(".docx"):
        import docx
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
# 問題の作成
# ---------------------------------------------------------------------------
def build_items(grade, selected, n, seed, use_random):
    """(items, notes) を返す。"""
    rng = random.Random(seed)
    notes = []
    chars = list(GRADE_KANJI[grade])
    if use_random:
        pool_k = rng.sample(chars, min(n, len(chars)))
    else:
        pool_k = list(selected)
        if len(pool_k) > n:
            pool_k = rng.sample(pool_k, n)
    kanji = "".join(pool_k)

    items = []
    if api_key():
        try:
            items = generate_with_claude(grade, kanji, n)
            if len(items) < n:
                notes.append(f"AIが作成できた問題は{len(items)}問でした。")
        except Exception as ex:
            notes.append(f"AIによる問題作成に失敗したため、サンプル問題を使用します（{ex}）")
    else:
        notes.append("ANTHROPIC_API_KEY が未設定のため、内蔵のサンプル問題から作成しました。")

    if len(items) < n:
        ks = set(pool_k)
        used = {e["word"] for e in items}
        allp = [e for e in sample_pool(grade) if e["word"] not in used]
        matched = [e for e in allp if any(c in ks for c in e["word"])]
        others = [e for e in allp if e not in matched]
        rng.shuffle(matched)
        rng.shuffle(others)
        need = n - len(items)
        if len(matched) < need and not use_random:
            notes.append("選んだ漢字を含む内蔵サンプルが少ないため、他のサンプル問題で補いました。")
        items += (matched + others)[:need]
    if len(items) < n:
        notes.append(f"問題数が足りず、{len(items)}問のプリントになりました。")
    return items[:n], notes


# ---------------------------------------------------------------------------
# プリントHTML生成（縦書き）
# ---------------------------------------------------------------------------
CSS = """
*{box-sizing:border-box}
html,body{margin:0}
body{background:#e9ecef;color:#111;
 font-family:"UD デジタル 教科書体 NK-R","UD Digi Kyokasho NK-R","Hiragino Mincho ProN","Yu Mincho","BIZ UDMincho","Noto Serif CJK JP","Noto Serif JP","IPAexMincho",serif}
.toolbar{position:sticky;top:0;z-index:5;display:flex;gap:8px;align-items:center;padding:8px;
 background:#fff;border-bottom:1px solid #ccc;font-family:sans-serif;font-size:13px;color:#555}
.toolbar button{font-size:16px;padding:10px 18px;min-height:44px;border:0;border-radius:8px;background:#1a73e8;color:#fff}
.scaler{transform-origin:top left;padding:8px}
.sheet{width:__W__mm;height:__H__mm;background:#fff;padding:5mm 6mm;margin:0 0 8px;display:flex;
 flex-direction:column;box-shadow:0 1px 6px #0003;overflow:hidden}
.body{flex:1;min-height:0;display:grid;direction:rtl;grid-template-columns:auto 1fr;gap:3mm}

/* 右端：テスト名／組・番／名前（1行）＋説明・得点 */
.names{direction:ltr;writing-mode:vertical-rl;height:100%}
.line{display:flex;inline-size:100%;margin-left:1.5mm}
.line1{block-size:11.5mm;flex-direction:row}
.line2{block-size:8mm;flex-direction:row}
.tt{flex:none;font-size:5mm;font-weight:bold;letter-spacing:.4mm;line-height:2.1;white-space:nowrap;padding-top:.5mm}
.tt.long{font-size:4mm}
.nf{display:flex;flex:none;border:.4mm solid #111;font-size:4.2mm;line-height:1.1}
.nf .lb{flex:none;text-align:center;padding:1.2mm 0;letter-spacing:.8mm}
.nf .sp{flex:1}
.nf.cls{block-size:11.5mm;inline-size:auto;height:36mm;margin-top:2mm}
.nf.name{block-size:11.5mm;flex:1;margin-top:2mm}
.desc{flex:1;font-size:3.2mm;line-height:1.7;color:#222}
.score{flex:none;border:.4mm solid #111;height:30mm;block-size:8mm;font-size:3.6mm;display:flex;margin-top:2mm}
.score span{padding:1mm 0;text-align:center;flex:none}
.score .sp{flex:1}

/* 問題エリア */
.qs{direction:rtl;display:grid;min-height:0;min-width:0;border-top:.25mm solid #888;border-right:.25mm solid #888}
.q{direction:ltr;writing-mode:vertical-rl;display:flex;flex-wrap:wrap;align-content:flex-start;align-items:center;
 padding:1.2mm 1mm;font-size:var(--fs);
 line-height:var(--lh);border-left:.25mm solid #888;border-bottom:.25mm solid #888;min-height:0;min-width:0}
.no{display:inline-block;width:5.4mm;height:5.4mm;border:.3mm solid #111;border-radius:50%;font-size:3.2mm;
 line-height:5mm;text-align:center;font-family:sans-serif;margin-inline-end:1.2mm;flex:none}
.u{flex:none}
.tcy{text-combine-upright:all;-webkit-text-combine:horizontal}
rt{font-size:.5em;line-height:1}

/* 解答欄：漢字ごとに□を連結／送り仮名つきは〔 〕枠。よみは枠の外（右）に配置 */
.answ{display:block;flex:none;position:relative;padding-right:calc(var(--fs)*.72);margin:.6mm 0}
.bxs{display:inline-flex}
.bx{display:block;width:var(--bx);height:var(--bx);border:.4mm solid #111;text-align:center;
 line-height:calc(var(--bx) - .8mm);font-size:calc(var(--bx)*.78);color:#c4001a;
 background:
  linear-gradient(90deg,transparent calc(50% - .1mm),#cfcfcf 0 calc(50% + .1mm),transparent 0),
  linear-gradient(transparent calc(50% - .1mm),#cfcfcf 0 calc(50% + .1mm),transparent 0)}
.bx+.bx{margin-top:-.4mm}
.brk{display:flex;align-items:center;justify-content:center;width:var(--bx);color:#c4001a;
 font-size:calc(var(--bx)*.78);letter-spacing:0;line-height:1;
 border-top:.45mm solid #111;border-bottom:.45mm solid #111;
 background:
  linear-gradient(#111,#111) 0 0/.45mm 2.6mm no-repeat,
  linear-gradient(#111,#111) 100% 0/.45mm 2.6mm no-repeat,
  linear-gradient(#111,#111) 0 100%/.45mm 2.6mm no-repeat,
  linear-gradient(#111,#111) 100% 100%/.45mm 2.6mm no-repeat}
.rd{position:absolute;right:0;top:0;bottom:0;display:flex;align-items:center;justify-content:center;
 font-size:calc(var(--fs)*.62);line-height:1;white-space:nowrap;color:#111}

/* お助けヒント */
.hint{display:block;flex:none;margin-top:.8mm;border:.3mm dashed #666;border-radius:1mm;padding:.5mm;vertical-align:top}
.hk{display:block;width:calc(var(--bx)*.6);height:calc(var(--bx)*.6);line-height:calc(var(--bx)*.6);
 font-size:calc(var(--bx)*.5);text-align:center;color:#444;
 -webkit-mask-image:linear-gradient(135deg,#000 0%,#000 40%,transparent 56%);
 mask-image:linear-gradient(135deg,#000 0%,#000 40%,transparent 56%)}

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


def tate(s: str) -> str:
    """縦書き用：半角数字を縦中横にする"""
    return re.sub(r"\d{1,3}", lambda m: f'<span class="tcy">{m.group()}</span>',
                  html.escape(s, quote=False))


def has_okurigana(word: str) -> bool:
    return any(not KANJI_RE.match(c) for c in word)


def answer_html(e, answer: bool) -> str:
    word = e["word"]
    rd = f'<span class="rd">{html.escape(e["reading"], quote=False)}</span>' if e.get("reading") else ""
    if has_okurigana(word):
        inner = html.escape(word, quote=False) if answer else ""
        return (f'<span class="answ"><span class="brk" style="height:calc(var(--bx)*{len(word)})">'
                f'{inner}</span>{rd}</span>')
    boxes = "".join(f'<span class="bx">{html.escape(c, quote=False) if answer else ""}</span>' for c in word)
    return f'<span class="answ"><span class="bxs">{boxes}</span>{rd}</span>'


def text_units(text: str, ruby: bool) -> str:
    """文を1文字（ルビ付き語は1まとまり）ずつの部品にする。折り返しを正確に制御するため。"""
    def chars(t):
        return "".join(f'<span class="u">{html.escape(c, quote=False)}</span>' for c in t)
    out, pos = [], 0
    for m in RUBY_RE.finditer(text):
        out.append(chars(text[pos:m.start()]))
        if ruby:
            out.append(f'<span class="u"><ruby>{html.escape(m.group(1), quote=False)}'
                       f'<rt>{html.escape(m.group(2), quote=False)}</rt></ruby></span>')
        else:
            out.append(chars(m.group(1)))
        pos = m.end()
    out.append(chars(text[pos:]))
    return "".join(out)


def sentence_html(e, ruby: bool, answer: bool) -> str:
    pre, sep, post = e["sentence"].partition("{W}")
    return text_units(pre, ruby) + answer_html(e, answer) + text_units(post, ruby)


def cell_html(i, e, row, col, help_on, ruby, answer) -> str:
    no = f'<span class="no"><span class="tcy">{i}</span></span>'
    hint = ""
    if help_on and not answer:
        hint = f'<span class="hint"><span class="hk">{html.escape(e["word"][0], quote=False)}</span></span>'
    return (f'<div class="q" style="grid-row:{row};grid-column:{col}">'
            f'{no}{sentence_html(e, ruby, answer)}{hint}</div>')


def layout_for(n):
    """(行数, 各行の問題数リスト, 解答枠サイズmm, 文字サイズmm)"""
    bands = 1 if n <= 6 else 2 if n <= 20 else 3 if n <= 30 else 4
    base, extra = divmod(n, bands)
    counts = [base + 1] * extra + [base] * (bands - extra)
    if n <= 10:
        bx, fs = 13, 5.2
    elif n <= 20:
        bx, fs = 12, 4.8
    elif n <= 30:
        bx, fs = 11.5, 4.4
    elif n <= 40:
        bx, fs = 11, 4.2
    else:
        bx, fs = 9.0, 3.5
    return bands, counts, bx, fs


def sheet_html(items, title, kind, answer, publisher, term) -> str:
    help_on = kind in ("お助け（2画ヒント）", "お助け＋ルビ")
    ruby = kind in ("ルビ付き", "お助け＋ルビ")
    n = len(items)
    bands, counts, bx, fs = layout_for(n)
    cols = max(counts)
    lh = 2.05 if ruby else 1.6
    pts = f"{100 / n:.1f}".rstrip("0").rstrip(".")

    cells, idx = [], 0
    for r, c in enumerate(counts, start=1):
        for k in range(1, c + 1):
            cells.append(cell_html(idx + 1, items[idx], r, k, help_on, ruby, answer))
            idx += 1

    ttl = title + ("（解答）" if answer else "（100点満点）")
    long = " long" if len(ttl) > 20 else ""
    desc = ("赤字が答えです。" if answer else
            f"文中の　線のひらがなを漢字で書きましょう。（1問{pts}点）")
    meta = f"{publisher}・{term}"
    names = (
        '<div class="names">'
        f'<div class="line line1"><div class="tt{long}">{tate(ttl)}</div>'
        '<div class="nf cls"><span class="sp"></span><span class="lb">組</span><span class="sp"></span>'
        '<span class="lb">番</span><span class="sp"></span></div>'
        '<div class="nf name"><span class="lb">名前</span><span class="sp"></span></div></div>'
        f'<div class="line line2"><div class="desc">{tate(meta)}　{tate(desc)}</div>'
        '<div class="score"><span>得点</span><span class="sp"></span></div></div>'
        '</div>'
    )
    return (f'<section class="sheet"><div class="body">{names}'
            f'<div class="qs" style="--fs:{fs}mm;--bx:{bx}mm;--lh:{lh};'
            f'grid-template-columns:repeat({cols},1fr);grid-template-rows:repeat({bands},1fr)">'
            f'{"".join(cells)}</div></div></section>')


def document_html(sheets, paper) -> str:
    w, h, size = PAPERS[paper]
    css = CSS.replace("__W__", str(w)).replace("__H__", str(h)).replace("__PAPER__", size)
    return (f'<!doctype html><html lang="ja"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1"><style>{css}</style></head><body>'
            f'<div class="toolbar"><button onclick="window.print()">🖨 共有・プリント</button>'
            f'<span>{paper}・横向きで印刷（倍率100%／余白なし）</span></div>'
            f'<div class="scaler">{"".join(sheets)}</div><script>{JS}</script></body></html>')


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------
def set_sel(key, chars):
    st.session_state[key] = list(chars)


def kanji_picker(grade):
    chars = list(GRADE_KANJI[grade])
    key = f"sel_{grade}"
    st.session_state.setdefault(key, [])
    b1, b2, b3 = st.columns([1, 1, 3])
    b1.button("全選択", key=f"all_{grade}", on_click=set_sel, args=(key, chars))
    b2.button("全解除", key=f"none_{grade}", on_click=set_sel, args=(key, []))
    if hasattr(st, "pills"):
        sel = st.pills("使う漢字", chars, selection_mode="multi", key=key, label_visibility="collapsed")
    else:
        sel = st.multiselect("使う漢字", chars, key=key, label_visibility="collapsed")
    sel = sel or []
    b3.caption(f"{grade}年生の漢字 {len(chars)}字のうち {len(sel)}字を選択中")
    return sel


def main():
    st.set_page_config(page_title="漢字小テスト作成", page_icon="📝", layout="wide")
    if not check_password():
        return
    st.markdown("<style>.stButton>button,.stDownloadButton>button{min-height:46px}"
                "div[role=radiogroup]{gap:.4rem}</style>", unsafe_allow_html=True)
    st.title("📝 漢字小テスト作成（縦書き）")

    st.session_state.setdefault("seed", random.randint(0, 10**6))

    with st.expander("⚙ 設定", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        grade = c1.selectbox("学年", [1, 2, 3, 4, 5, 6], format_func=lambda g: f"{g}年生")
        publisher = c2.selectbox("教科書会社", PUBLISHERS)
        term = c3.selectbox("学期", TERMS)
        paper = c4.selectbox("用紙", list(PAPERS))
        c5, c6 = st.columns([1, 3])
        n = c5.selectbox("問題数", COUNTS, index=1, format_func=lambda x: f"{x}問")
        title = c6.text_input("プリント名", value=f"{grade}年 漢字小テスト")
        kind = st.radio("プリントの種類", KINDS, horizontal=True)

    with st.expander("🔤 使う漢字を選ぶ（学年別漢字配当表）", expanded=True):
        how = st.radio("漢字の選び方", ["一覧から選ぶ", "学年の漢字からランダムに選ぶ"], horizontal=True)
        selected = kanji_picker(grade) if how == "一覧から選ぶ" else []
        if how != "一覧から選ぶ":
            st.caption(f"{grade}年生の漢字 {len(GRADE_KANJI[grade])}字から、問題数に合わせてランダムに選びます。")

    with st.expander("📎 既存プリントを添付（任意）"):
        up = st.file_uploader("txt / csv / docx / 画像 / PDF（問題をそのまま使います）",
                              type=["txt", "csv", "md", "docx", "png", "jpg", "jpeg", "pdf"])
        st.caption("txt・csv・docx は「【語|よみ】を含む文」または「語,よみ,例文」の形式で読み取ります。"
                   "画像・PDFは AI が読み取ります（APIキーが必要）。")

    if st.button("✨ プリントを作成", type="primary"):
        st.session_state["seed"] = random.randint(0, 10**6)
        items, notes = None, []
        if up is not None:
            try:
                att = read_attachment(up)
                if att:
                    items = att[:n]
                    notes.append(f"添付プリントから {len(att)} 問を読み取りました。")
                else:
                    notes.append("添付プリントから問題を読み取れませんでした。")
            except Exception as ex:
                notes.append(f"添付ファイルの読み取りに失敗しました: {ex}")
        if items is None:
            if how == "一覧から選ぶ" and not selected:
                st.warning("使う漢字を1字以上選んでください。")
            else:
                with st.spinner("問題を作成中…（30秒〜1分ほどかかることがあります）"):
                    items, nn = build_items(grade, selected, n, st.session_state["seed"],
                                            how != "一覧から選ぶ")
                notes += nn
        if items:
            st.session_state["items"] = items
        st.session_state["notes"] = notes

    for msg in st.session_state.get("notes", []):
        st.info(msg)

    items = st.session_state.get("items")
    if not items:
        st.caption("設定と漢字を選んで「プリントを作成」を押してください。")
        return

    q_html = sheet_html(items, title, kind, False, publisher, term)
    a_html = sheet_html(items, title, kind, True, publisher, term)
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


if __name__ == "__main__":
    main()

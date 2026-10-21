"""小学生向け 漢字テスト作成アプリ (Streamlit + Claude API)

起動:
    pip install streamlit anthropic
    streamlit run app.py

設定 (.streamlit/secrets.toml または環境変数)
    APP_PASSWORD      = "amida"
    ANTHROPIC_API_KEY = "sk-ant-..."
"""

import hmac
import json
import os
import re

import anthropic
import streamlit as st

MODEL = "claude-sonnet-5-5"

PUBLISHERS = ["東京書籍", "光村図書", "教育出版", "学校図書", "三省堂", "その他（指定なし）"]

# 学年別 新出漢字（目安）。必ず文部科学省の学年別漢字配当表と照合し、
# 画面上のテキスト欄で自由に編集できます。5・6年は空欄なので貼り付けてください。
KANJI_BY_GRADE = {
    1: "一右雨円王音下火花貝学気九休玉金空月犬見五口校左三山子四糸字耳七車手十出女小上森人水正生青夕石赤千川先早草足村大男竹中虫町天田土二日入年白八百文木本名目立力林六",
    2: "引羽雲園遠何科夏家歌画回会海絵外角楽活間丸岩顔汽記帰弓牛魚京強教近兄形計元言原戸古午後語工公広交光考行高黄合谷国黒今才細作算止市矢姉思紙寺自時室社弱首秋週春書少場色食心新親図数西声星晴切雪船線前組走多太体台地池知茶昼長鳥朝直通弟店点電刀冬当東答頭同道読内南肉馬売買麦半番父風分聞米歩母方北毎妹万明鳴毛門夜野友用曜来里理話",
    3: "悪安暗医委意育員院飲運泳駅央横屋温化荷界開階寒感漢館岸起期客究急級宮球去橋業曲局銀区苦具君係軽血決研県庫湖向幸港号根祭皿仕死使始指歯詩次事持式実写者主守取酒受州拾終習集住重宿所暑助昭消商章勝乗植申身神真深進世整昔全相送想息速族他打対待代第題炭短談着注柱丁帳調追定庭笛鉄転都度投豆島湯登等動童農波配倍箱畑発反坂板皮悲美鼻筆氷表秒病品負部服福物平返勉放味命面問役薬由油有遊予羊洋葉陽様落流旅両緑礼列練路和",
    4: "愛案以衣位茨印英栄媛塩岡億加果貨課芽賀改械害街各覚潟完官管関観願岐希季旗器機議求泣給挙漁共協鏡競極熊訓軍郡径型景芸欠結建健験固功好香候康佐差菜最埼材崎昨札刷察参産散残氏司試児治滋辞鹿失借種周祝順初松笑唱焼象照賞臣信成省清静席積折節説浅戦選然争倉巣束側続卒孫帯隊達単置仲貯兆腸低底停的典伝徒努灯堂働特得毒熱念敗梅博飯飛費必票標不夫付府副粉兵別辺変便包法望牧末満未脈民無約勇要養浴利陸良料量輪類令冷例歴連老労録",
    5: "圧囲移因永営演応往桜可仮価河過賀快解格確額刊幹慣眼紀軌窮許可局銀句苦軍群径茎劇結建個候航校溝鉱講貢購載細佐査再災妻採際在財罪雑酸賛支志枝師資飼示似識質舎謝授修述術準序招承証条状常情織職制性政勢精製税責績接設舌絶銭祖素総造像増則測属率貸損退担任探談値中柱庭停的典伝殿島湯登塗奴怒倒党冬統導働胴登録独読縄難任燃能破犯販評貧布婦富敷部複復標秒仏粉兵別辺返勉弁保墓貿易報豊防貿暴務夢迷綿輸余預容略領輪類",
    6: "異遺域宇映延沿我灰拡革閣割株干巻看簡危机揮貴疑吸供胸郷勤筋系敬警劇激穴絹権憲源厳己呼誤后孝皇紅降鋼刻穀骨困砂座済裁策冊蚕至私姿視詞誌磁射捨尺若樹収宗就衆従縦縮熟純処署諸除将傷障城蒸針仁垂推寸盛聖誠宣専泉洗染善奏創装層操蔵臓存尊宅担探誕段暖値宙忠著庁頂潮賃痛展討党糖届難乳認納脳派拝背肺俳班晩否批秘腹奮並陛閉片補暮宝訪亡忘棒枚幕描木郵優猶裕幼欲翌乱卵覧裏律臨朗論和",
}

SYSTEM_PROMPT = """あなたは日本の小学校の経験豊富な教師であり、教科書会社の漢字テストを作る専門家です。
業者テスト（市販の単元テスト）のように、教科書の単元の文脈に沿った自然な短文で問題を作ります。
出力は指定されたJSONのみとし、前後に説明文やコードブロック記号を付けないでください。"""


# ---------- 認証 ----------
def check_password() -> bool:
    """簡易パスワード認証。通れば True。"""
    if st.session_state.get("authenticated"):
        return True

    expected = st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", ""))
    st.title("🔒 漢字テスト作成アプリ")

    if not expected:
        st.error("パスワードが未設定です。APP_PASSWORD を secrets.toml か環境変数に設定してください。")
        return False

    with st.form("login"):
        pw = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン")
    if submitted:
        if hmac.compare_digest(pw.encode(), expected.encode()):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが違います。")
    return False


# ---------- 生成 ----------
def build_prompt(grade, publisher, unit, n_read, n_write, kanji) -> str:
    return f"""次の条件で小学生の漢字テストを作成してください。

【条件】
- 学年: 小学{grade}年生
- 教科書会社: {publisher}
- 単元名: {unit or "指定なし"}
- 読み問題: {n_read}問 / 書き問題: {n_write}問
- 出題対象の新出漢字: {kanji}

【作成ルール】
1. 出題する漢字（読み問題の太字部分・書き問題の答え）は、上記の「出題対象の新出漢字」から選ぶ。できるだけ重複せず、全体にばらけさせる。
2. 問題文は{grade}年生が読める文にする。問題文中の漢字は小学{grade}年生までに習う漢字だけを使い、それ以外はひらがなにする。
3. 単元名や教科書会社の傾向を意識し、単元に関連する語彙・場面の短文にする。ただし教科書本文の引用はしない。
4. 読み問題: 文中の対象の語を【】で囲む。解答は【】内の読み（ひらがな）。
5. 書き問題: 文中の対象の語をひらがな（または送り仮名以外ひらがな）にして【】で囲む。解答は漢字（送り仮名含む）。
6. 熟語・送り仮名のある語・訓読み・音読みをバランスよく混ぜる。
7. 難易度は易しい問題から徐々に難しくする。

【出力JSON形式】
{{
  "title": "テストのタイトル",
  "read": [{{"sentence": "【朝顔】が咲いた。", "answer": "あさがお"}}],
  "write": [{{"sentence": "【がっこう】へ行く。", "answer": "学校"}}]
}}
read は {n_read} 件、write は {n_write} 件ちょうど出力すること。"""


def generate_test(prompt: str) -> dict:
    api_key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", ""))
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY が設定されていません。")
    client = anthropic.Anthropic(api_key=api_key)
    resp = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        raise ValueError("AIの応答からJSONを取得できませんでした。もう一度お試しください。")
    return json.loads(m.group(0))


# ---------- 表示 ----------
def to_text(test: dict, grade, publisher, unit) -> str:
    lines = [test.get("title", "漢字テスト"), f"小学{grade}年 {publisher} {unit}",
             "名前（　　　　　　　　）　　　　　点（　　／　　）", ""]
    n = 1
    if test["read"]:
        lines.append("■ 次の　　線の漢字の読みを、ひらがなで書きましょう。")
        for q in test["read"]:
            lines.append(f"({n}) {q['sentence']}")
            n += 1
        lines.append("")
    if test["write"]:
        lines.append("■ 次の　　線のひらがなを、漢字（送りがなも）で書きましょう。")
        for q in test["write"]:
            lines.append(f"({n}) {q['sentence']}")
            n += 1
        lines.append("")
    lines.append("【解答】")
    n = 1
    for q in test["read"] + test["write"]:
        lines.append(f"({n}) {q['answer']}")
        n += 1
    return "\n".join(lines)


def show_test(test: dict, grade, publisher, unit):
    st.header(test.get("title", "漢字テスト"))
    st.caption(f"小学{grade}年 / {publisher} / {unit or '単元指定なし'}")
    st.write("名前（　　　　　　　　）　　点（　　／　　）")

    n = 1
    if test["read"]:
        st.subheader("一　読み問題")
        st.write("次の【　】の漢字の読みを、ひらがなで書きましょう。")
        for q in test["read"]:
            st.markdown(f"**({n})** {q['sentence']}")
            n += 1
    if test["write"]:
        st.subheader("二　書き問題")
        st.write("次の【　】のひらがなを、漢字（送りがなも）で書きましょう。")
        for q in test["write"]:
            st.markdown(f"**({n})** {q['sentence']}")
            n += 1

    with st.expander("✅ 解答を表示"):
        n = 1
        for q in test["read"]:
            st.markdown(f"**({n})** {q['answer']}　（{q['sentence']}）")
            n += 1
        for q in test["write"]:
            st.markdown(f"**({n})** {q['answer']}　（{q['sentence']}）")
            n += 1

    st.download_button("📥 テキストでダウンロード", to_text(test, grade, publisher, unit),
                       file_name="kanji_test.txt")


# ---------- メイン ----------
def main():
    st.set_page_config(page_title="漢字テスト作成", page_icon="📝", layout="wide")
    if not check_password():
        return

    st.title("📝 小学生 漢字テスト作成")

    with st.sidebar:
        st.header("条件設定")
        grade = st.selectbox("学年", [1, 2, 3, 4, 5, 6], format_func=lambda g: f"{g}年生")
        publisher = st.selectbox("教科書会社", PUBLISHERS)
        unit = st.text_input("単元名", placeholder="例：ふしぎ大図鑑")
        n_total = st.slider("問題数", 10, 50, 20)
        read_ratio = st.slider("読み問題の割合（%）", 0, 100, 50, step=10)
        st.caption(f"書き問題の割合：{100 - read_ratio}%")
        if st.button("ログアウト"):
            st.session_state["authenticated"] = False
            st.rerun()

    kanji = st.text_area(
        f"{grade}年生で習う新出漢字（出題範囲・編集可）",
        value=KANJI_BY_GRADE.get(grade, ""),
        key=f"kanji_{grade}",
        height=140,
        help="単元で習う漢字だけに絞ると、より業者テストらしくなります。",
    )
    kanji = "".join(ch for ch in kanji if not ch.isspace())

    n_read = round(n_total * read_ratio / 100)
    n_write = n_total - n_read
    st.info(f"読み {n_read}問 ＋ 書き {n_write}問 ＝ 全{n_total}問　／　出題範囲 {len(kanji)}字")

    if st.button("🚀 テストを生成", type="primary"):
        if not kanji:
            st.warning("出題範囲の漢字を入力してください。")
        else:
            with st.spinner("テストを作成中…（30秒〜1分ほどかかります）"):
                try:
                    test = generate_test(build_prompt(grade, publisher, unit, n_read, n_write, kanji))
                    test.setdefault("read", [])
                    test.setdefault("write", [])
                    st.session_state["test"] = (test, grade, publisher, unit)
                except Exception as e:
                    st.error(f"生成に失敗しました: {e}")

    if "test" in st.session_state:
        show_test(*st.session_state["test"])


main()

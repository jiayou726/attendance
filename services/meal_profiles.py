"""供餐對象的午餐熱量基準。

數值全部照抄官方公告文件，不自行推估：

1. 國小、國中、高中
   教育部「學校午餐食物內容及營養基準」109 年 12 月 28 日修訂版。
   文件第 1 頁「一、學校午餐營養建議量」的熱量(大卡)欄位。

2. 幼兒園
   教育部「幼兒園餐點食物內容及營養基準」112 年 3 月 27 日臺教授國部字
   第 1120029103A 號令訂定。文件列出的 550／700 大卡是「早點＋午餐＋午點」
   合計（占每日 4/9）。依同頁註 2 的計算方式，一天分三時段各占 1/3，
   點心熱量為正餐的 1/2，因此早點 1/9、午餐 2/9、午點 1/9；
   午餐 = 2/9 ÷ 4/9 = 合計值的一半，即 275 與 350 大卡。
   幼兒園公告的是單一建議值而非區間，這裡以 ±10% 作為排菜可接受範圍，
   並在備註寫明這是排菜用的容許範圍，不是官方區間。

來源連結與版本日期見 docs/nutrition_sources.md。
"""

from __future__ import annotations

from decimal import Decimal

SCHOOL_DOC = "教育部 學校午餐食物內容及營養基準"
SCHOOL_DOC_VERSION = "109.12.28 修訂"
SCHOOL_DOC_URL = (
    "https://mlunch.nat.gov.tw/manasystem/files/law/"
    "1100412135351_學校午餐營養基準-109年修正.pdf"
)

PRESCHOOL_DOC = "教育部 幼兒園餐點食物內容及營養基準"
PRESCHOOL_DOC_VERSION = "112.03.27 訂定（臺教授國部字第1120029103A號令）"
PRESCHOOL_DOC_URL = (
    "https://www.ece.moe.edu.tw/ch/filelist/preschool_education/"
    "preschool_education_20230327/"
)

PRESCHOOL_NOTE_TEMPLATE = (
    "官方公告 {combined} 大卡為早點＋午餐＋午點合計（占每日 4/9）；"
    "依註 2 的時段拆分，午餐占 2/9，即 {lunch} 大卡。"
    "官方未給區間，此處 ±10% 為排菜容許範圍，非官方數值。"
)

# code, name, kcal_min, kcal_max, sort_order, document, version, url, note
PROFILE_SEEDS: tuple[tuple, ...] = (
    (
        "preschool_2_3", "幼兒園 2-3 歲", 248, 303, 10,
        PRESCHOOL_DOC, PRESCHOOL_DOC_VERSION, PRESCHOOL_DOC_URL,
        PRESCHOOL_NOTE_TEMPLATE.format(combined=550, lunch=275),
    ),
    (
        "preschool_4_6", "幼兒園 4-6 歲", 315, 385, 20,
        PRESCHOOL_DOC, PRESCHOOL_DOC_VERSION, PRESCHOOL_DOC_URL,
        PRESCHOOL_NOTE_TEMPLATE.format(combined=700, lunch=350),
    ),
    (
        "elementary_1_3", "國小低年級（1-3 年級）", 620, 720, 30,
        SCHOOL_DOC, SCHOOL_DOC_VERSION, SCHOOL_DOC_URL,
        "官方午餐熱量建議量 620~720 大卡。",
    ),
    (
        "elementary_4_6", "國小高年級（4-6 年級）", 720, 830, 40,
        SCHOOL_DOC, SCHOOL_DOC_VERSION, SCHOOL_DOC_URL,
        "官方午餐熱量建議量 720~830 大卡。",
    ),
    (
        "junior_high", "國中", 800, 930, 50,
        SCHOOL_DOC, SCHOOL_DOC_VERSION, SCHOOL_DOC_URL,
        "官方午餐熱量建議量 800~930 大卡。",
    ),
    (
        "senior_high_male", "高中（男生）", 900, 1050, 60,
        SCHOOL_DOC, SCHOOL_DOC_VERSION, SCHOOL_DOC_URL,
        "官方午餐熱量建議量 900~1050 大卡。",
    ),
    (
        "senior_high_female", "高中（女生）", 680, 810, 70,
        SCHOOL_DOC, SCHOOL_DOC_VERSION, SCHOOL_DOC_URL,
        "官方午餐熱量建議量 680~810 大卡。",
    ),
    (
        "senior_high_mixed", "高中（男女混合）", 680, 1050, 80,
        SCHOOL_DOC, SCHOOL_DOC_VERSION, SCHOOL_DOC_URL,
        "男女生班級混合供餐時，取男生 900~1050 與女生 680~810 的聯集，"
        "實務上仍建議依實際班級性別比另行檢視。",
    ),
)


def seed_profiles(session) -> int:
    """建立或更新供餐對象基準。回傳有變動的筆數。

    使用者自行改過的數值不會被覆寫；只有和官方公告不一致
    且來源仍標記為本模組的文件時才會更新，避免每次啟動就蓋掉人工調整。
    """

    from models import KitchenMealProfile

    changed = 0
    for (code, name, kcal_min, kcal_max, sort_order,
         document, version, url, note) in PROFILE_SEEDS:
        profile = KitchenMealProfile.query.filter_by(code=code).one_or_none()
        if profile is None:
            profile = KitchenMealProfile(code=code)
            session.add(profile)
            changed += 1
        elif (
            profile.kcal_min == Decimal(kcal_min)
            and profile.kcal_max == Decimal(kcal_max)
            and profile.name == name
            and profile.source_version == version
        ):
            continue
        else:
            changed += 1
        profile.name = name
        profile.meal_type = "午餐"
        profile.kcal_min = Decimal(kcal_min)
        profile.kcal_max = Decimal(kcal_max)
        profile.source_document = document
        profile.source_version = version
        profile.source_url = url
        profile.note = note
        profile.sort_order = sort_order
        profile.active = True
    return changed


def active_profiles():
    from models import KitchenMealProfile

    return (
        KitchenMealProfile.query
        .filter(KitchenMealProfile.active.is_(True))
        .order_by(KitchenMealProfile.sort_order, KitchenMealProfile.id)
        .all()
    )

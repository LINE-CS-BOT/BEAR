from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    # LINE Messaging API
    LINE_CHANNEL_ACCESS_TOKEN: str
    LINE_CHANNEL_SECRET: str
    LINE_GROUP_ID: str = ""           # 內部人員群組 ID（bot 無法回答時通知真人）
    LINE_GROUP_ID_HQ: str = ""        # 總公司群組 ID（庫存不足時詢問調貨）
    LINE_GROUP_ID_SHOWCASE: str = ""  # 看貨群 ID（只回營業時間，其餘靜默）
    ADMIN_LINE_UID: str = "Uac17599b38b673b836ccb48025204b19"   # 小熊本人（系統異常/狀態報告）
    HELPER_LINE_UID: str = "U8664e671a26d2eca1237fe94ad634205"  # 小蠻牛-新北小幫手（人工客服動作）

    # Ecount ERP API
    ECOUNT_COMPANY_NO: str = ""
    ECOUNT_USER_ID: str = ""
    ECOUNT_API_CERT_KEY: str = ""
    ECOUNT_ZONE: str = "IB"
    ECOUNT_BASE_URL: str = "https://oapiIB.ecount.com"
    ECOUNT_DEFAULT_CUST_CD: str = "LINECUST"  # LINE 新客戶預設客戶代碼

    # 總公司 Ecount（上游供應商，採購單流程）
    HQ_ECOUNT_COMPANY_NO: str = ""
    HQ_ECOUNT_USER_ID: str = ""
    HQ_ECOUNT_API_CERT_KEY: str = ""
    HQ_ECOUNT_WEB_PASSWORD: str = ""
    HQ_ECOUNT_ZONE: str = "IB"
    HQ_ECOUNT_BASE_URL: str = "https://oapiIB.ecount.com"
    HQ_ECOUNT_OUR_CUST_CD: str = "A05"        # 小蠻牛在總公司端的客戶編碼
    HQ_ECOUNT_OUR_WH_CD: str = "200"          # 總公司視角我方收貨倉庫

    # Google Calendar
    GOOGLE_CALENDAR_ID: str = ""
    GOOGLE_CREDENTIALS_FILE: str = "credentials.json"
    GOOGLE_TOKEN_FILE: str = "token.json"

    # 營業時間
    BUSINESS_HOURS_START: str = "13:00"
    BUSINESS_HOURS_END: str = "21:00"
    BUSINESS_DAYS: str = "2,3,4,5,6,7"  # 逗號分隔，1=週一，週一公休
    BUSINESS_TZ: str = "Asia/Taipei"
    STORE_ADDRESS: str = "新北市土城區中央路二段394巷12號"

    # Admin 介面登入（HTTP Basic Auth）
    ADMIN_USER: str = "admin"
    ADMIN_PASS: str = "changeme"

    # 產品媒體路徑（圖片 + 影片）
    PRODUCT_MEDIA_PATH: str = r"H:\其他電腦\我的電腦\小蠻牛\產品照片"

    # 聯絡群組清單（逗號分隔，「聯絡群組 Z3456」指令推送目標）
    CONTACT_GROUP_CHATS: str = (
        "VIP丞&Wei的聯絡群組,"
        "VIP程品勝的聯絡群組,"
        "VIP翁敬恩的聯絡群組,"
        "林子翔的樹林聯絡群組"
    )

    def contact_group_chats_list(self) -> list[str]:
        return [s.strip() for s in self.CONTACT_GROUP_CHATS.split(",") if s.strip()]

    def business_days_list(self) -> list[int]:
        return [int(d.strip()) for d in self.BUSINESS_DAYS.split(",")]


settings = Settings()

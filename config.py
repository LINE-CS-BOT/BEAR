from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    # LINE Messaging API
    LINE_CHANNEL_ACCESS_TOKEN: str
    LINE_CHANNEL_SECRET: str
    LINE_GROUP_ID: str = ""           # 內部人員群組 ID（bot 無法回答時通知真人）
    LINE_GROUP_ID_HQ: str = ""        # 總公司群組 ID（庫存不足時詢問調貨）
    LINE_GROUP_ID_SHOWCASE: str = ""  # 看貨群 ID（只回營業時間，其餘靜默）

    # Ecount ERP API
    ECOUNT_COMPANY_NO: str = ""
    ECOUNT_USER_ID: str = ""
    ECOUNT_API_CERT_KEY: str = ""
    ECOUNT_ZONE: str = "IB"
    ECOUNT_BASE_URL: str = "https://oapiIB.ecount.com"
    ECOUNT_DEFAULT_CUST_CD: str = "LINECUST"  # LINE 新客戶預設客戶代碼

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

    def business_days_list(self) -> list[int]:
        return [int(d.strip()) for d in self.BUSINESS_DAYS.split(",")]


settings = Settings()

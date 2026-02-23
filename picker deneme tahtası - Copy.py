import os
import ssl
import urllib3
import requests
import time
import re
import imaplib
import email
import gspread
import traceback
from datetime import datetime
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# ==========================================
# 🛑 SSL VE GÜVENLİK YAMASI
# ==========================================
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except AttributeError:
    pass

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_original_request = requests.Session.request
def _patched_request(self, method, url, *args, **kwargs):
    kwargs['verify'] = False
    return _original_request(self, method, url, *args, **kwargs)
requests.Session.request = _patched_request

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
os.environ['CURL_CA_BUNDLE'] = ''
os.environ['REQUESTS_CA_BUNDLE'] = ''

# ==========================================
# ⚙️ AYARLAR
# ==========================================
GMAIL_ADRES = "fabur.rpa.ext@yemeksepeti.com"
GMAIL_APP_PASSWORD = "icru ofkn askl xrrc"
OKTA_USERNAME = "fabür.rpa.ext"
OKTA_PASSWORD = "Banabi22072019*"

ROOSTER_URL = "https://dp-tr.eu.logisticsbackoffice.com/dashboard/rooster/workers?filter_status=active_contract&page=1&size=10"
CARSI_URL   = "https://carsi-portal.yemeksepeti.com/pv2/tr/p/picker-performance/accounts?businessUnit=DMART"

DATA_SHEET_URL = "https://docs.google.com/spreadsheets/d/1xZBNdH25rCxVay43jBgIsNorKM31-73rA5du3XZWn34/edit"
DATA_TAB_NAME  = "Nazlıca"
MAP_SHEET_URL  = "https://docs.google.com/spreadsheets/d/1PN3daa9H-Odd7Gdpuiuef8U85ko4Hb5DAtxETQDXWaA/edit"
MAP_TAB_NAME   = "picker"

# Google Sheets sütun numaraları (1-indexed)
COL_TCKN      = 1   # A
COL_NAME      = 2   # B
COL_SURNAME   = 3   # C
COL_DEPOT     = 7   # G
COL_PHONE     = 8   # H
COL_EMAIL     = 11  # K
COL_STATUS    = 15  # O
COL_EMP_ID    = 16  # P
COL_PORTAL_ST = 17  # Q

# ==========================================
# 🔧 YARDIMCI FONKSİYONLAR
# ==========================================

def tr_char_replace(text):
    if not text:
        return ""
    tr_map = {
        'ç': 'c', 'Ç': 'c', 'ğ': 'g', 'Ğ': 'g',
        'ı': 'i', 'I': 'i', 'İ': 'i',
        'ö': 'o', 'Ö': 'o', 'ş': 's', 'Ş': 's',
        'ü': 'u', 'Ü': 'u'
    }
    for tr, en in tr_map.items():
        text = text.replace(tr, en)
    return text.lower().strip()


def safe_col(row, col):
    return row[col - 1].strip() if len(row) >= col else ""


def get_sheets_client():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    JSON_KEY = os.path.join(BASE_DIR, "service_account.json")
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(JSON_KEY, scopes=scope)
    return gspread.authorize(creds)


def get_latest_okta_otp(timeout=90):
    print("  📩 Gmail'e bağlanıyor, OTP bekleniyor...")
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        mail = imaplib.IMAP4_SSL("imap.gmail.com", ssl_context=ctx)
        mail.login(GMAIL_ADRES, GMAIL_APP_PASSWORD)
        mail.select("inbox")

        deadline = time.time() + timeout
        while time.time() < deadline:
            status, data = mail.search(None, '(FROM "no-reply@okta.deliveryhero.com")')
            if status == 'OK' and data[0]:
                latest_id = data[0].split()[-1]
                _, msg_data = mail.fetch(latest_id, "(RFC822)")
                msg = email.message_from_bytes(msg_data[0][1])
                text = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            text += part.get_payload(decode=True).decode("utf-8", errors="ignore")
                else:
                    text = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
                match = re.search(r"\b\d{6}\b", text)
                if match:
                    code = match.group(0)
                    print(f"  🎯 OTP alındı: {code}")
                    return code
            time.sleep(5)
    except Exception as e:
        print(f"  ❌ OTP hatası: {e}")
    return None


def login_with_okta(driver):
    print("🔐 Okta girişi başlıyor...")
    wait = WebDriverWait(driver, 25)
    try:
        wait.until(EC.presence_of_element_located((By.NAME, "identifier"))).send_keys(OKTA_USERNAME)
        driver.find_element(By.CSS_SELECTOR, "input.button.button-primary").click()
        time.sleep(3)

        wait.until(EC.presence_of_element_located((By.NAME, "credentials.passcode"))).send_keys(OKTA_PASSWORD)
        driver.find_element(By.CSS_SELECTOR, "input.button.button-primary").click()
        time.sleep(8)

        if "email" in driver.page_source.lower() or "verify" in driver.current_url.lower():
            try:
                driver.find_element(By.XPATH, "//input[@type='submit' and @value='Send me an email']").click()
                print("  → 'Send me an email' tıklandı, 15 sn bekleniyor...")
                time.sleep(15)
            except Exception:
                print("  → OTP butonu bulunamadı, otomatik gönderilmiş olabilir.")

            otp = get_latest_okta_otp()
            if not otp:
                print("  ❌ OTP alınamadı!")
                return False

            driver.execute_script(
                f"var el=document.querySelector('input[name=\"credentials.passcode\"]');"
                f"el.value='{otp}';"
                f"el.dispatchEvent(new Event('input',{{bubbles:true}}));"
            )
            time.sleep(2)
            driver.find_element(By.XPATH, "//input[@value='Verify' or @type='submit']").click()
            print("  ✅ Verify tıklandı, yükleniyor...")
            time.sleep(10)

        return True
    except Exception as e:
        print(f"  ❌ Giriş hatası: {e}")
        return False


def dismiss_popups(driver, wait, timeout=5):
    print("  🔍 Pop-up kontrol ediliyor...")
    closed_any = False

    close_texts = ["Close", "Dismiss", "Got it", "OK", "Tamam", "Skip", "Cancel", "No thanks", "×", "✕"]
    for text in close_texts:
        try:
            btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH,
                f"//button[normalize-space(text())='{text}'] "
                f"| //button[contains(normalize-space(text()),'{text}')] "
                f"| //span[normalize-space(text())='{text}']/ancestor::button "
                f"| //a[normalize-space(text())='{text}']"
            )))
            driver.execute_script("arguments[0].click();", btn)
            print(f"  ✅ Pop-up kapatıldı (buton: '{text}')")
            closed_any = True
            time.sleep(1)
            break
        except Exception:
            continue

    icon_xpaths = [
        "//*[@data-test='close-button']", "//*[@data-test='dismiss-button']",
        "//*[@aria-label='Close']", "//*[@aria-label='Dismiss']",
        "//button[contains(@class,'close')]", "//button[contains(@class,'dismiss')]",
        "//div[contains(@class,'modal')]//button[contains(@class,'close')]",
        "//div[contains(@class,'popup')]//button", "//div[contains(@class,'toast')]//button",
    ]
    for xpath in icon_xpaths:
        try:
            btn = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].click();", btn)
            print(f"  ✅ Pop-up kapatıldı.")
            closed_any = True
            time.sleep(1)
            break
        except Exception:
            continue

    try:
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
    except Exception:
        pass

    if not closed_any:
        print("  ℹ️ Aktif pop-up bulunamadı.")


def ensure_logged_in(driver, target_url):
    driver.get(target_url)
    time.sleep(5)
    if "okta" in driver.current_url or "login" in driver.current_url:
        return login_with_okta(driver)
    print("  ✅ Oturum zaten açık.")
    return True


# ==========================================
# 🔵 FAZ 1: ROOSTER
# ==========================================

def rooster_create_worker(driver, wait, row, mapping_dict):
    tckn       = safe_col(row, COL_TCKN)
    raw_name   = safe_col(row, COL_NAME)
    surname    = safe_col(row, COL_SURNAME)
    depot_code = tr_char_replace(safe_col(row, COL_DEPOT))
    phone      = safe_col(row, COL_PHONE)
    email_addr = safe_col(row, COL_EMAIL)

    print(f"\n  👤 Rooster: {raw_name} {surname} işleniyor...")

    try:
        driver.get(ROOSTER_URL)
        time.sleep(5)

        create_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "Create worker")]')))
        create_btn.click()
        time.sleep(3)

        wait.until(EC.presence_of_element_located((By.XPATH, '//input[@id="InputControl_name"]'))).send_keys(f"{raw_name} {surname}")
        driver.find_element(By.XPATH, '//input[@id="InputControl_email"]').send_keys(email_addr)

        clean_phone = re.sub(r'[\s()\-]', '', phone)
        if clean_phone.startswith("0"):
            clean_phone = clean_phone[1:]
        driver.find_element(By.XPATH, '//input[@type="tel"]').send_keys(clean_phone)
        driver.find_element(By.XPATH, '//input[@name="password" or @type="password"]').send_keys("12345678@Aa")

        try:
            id_input = driver.find_element(By.XPATH, '//input[contains(@id,"id_number") or @placeholder="Paper No..."]')
            if not id_input.is_displayed():
                driver.find_element(By.XPATH, '//h3[contains(text(),"Identification")]').click()
                time.sleep(1)
            id_input.send_keys(tckn)
        except Exception:
            pass

        time.sleep(2)
        driver.find_element(By.XPATH, '//button[contains(.,"Save Details") or contains(.,"Save")]').click()
        print("  💾 Kaydet tıklandı, Employee ID bekleniyor...")
        time.sleep(5)

        employee_id = ""
        try:
            emp_el = wait.until(EC.visibility_of_element_located((By.XPATH, '//input[@id="id"]')))
            employee_id = emp_el.get_attribute('value') or ""
            if employee_id:
                print(f"  🆔 Employee ID: {employee_id}")
        except Exception as e:
            print(f"  ⚠️ Employee ID yakalanamadı: {e}")

        _rooster_create_contract(driver, wait, mapping_dict, depot_code)
        _rooster_assign_depot(driver, wait, mapping_dict, depot_code)

        return True, employee_id

    except Exception as e:
        return False, str(e)[:120]


def _rooster_create_contract(driver, wait, mapping_dict, depot_code):
    print("  📄 Sözleşme oluşturuluyor...")
    try:
        contract_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[.//span[text()="Contract Info"]]')))
        driver.execute_script("arguments[0].click();", contract_tab)
        time.sleep(4)

        wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@data-test="create-new-contract"]'))).click()
        time.sleep(5)

        try:
            wrapper = wait.until(EC.element_to_be_clickable((By.XPATH,
                "//div[contains(text(),'Contract name')]/ancestor::div[contains(@class,'Select-control')]"
                " | (//div[contains(@class,'Select-control')])[1]"
            )))
            wrapper.click()
            time.sleep(1)
            opt = wait.until(EC.element_to_be_clickable((By.XPATH,
                "//div[contains(@class,'Select-option') and normalize-space(text())='In house']"
                " | //div[normalize-space(text())='In house']"
            )))
            opt.click()
            print("  ✅ Contract Name: In house")
        except Exception as e:
            print(f"  ⚠️ Contract Name fallback: {e}")
            webdriver.ActionChains(driver).send_keys("In house", Keys.ENTER).perform()
        time.sleep(2)

        try:
            depot_data  = mapping_dict.get(depot_code, {})
            city_full   = depot_data.get("city", "Istanbul")
            city_search = city_full[1:] if len(city_full) > 1 else city_full
            city_ctrl = wait.until(EC.element_to_be_clickable((By.XPATH,
                "//label[contains(text(),'City')]/following-sibling::div//div[contains(@class,'Select-control')]"
                " | (//div[contains(@class,'Select-control')])[2]"
            )))
            city_ctrl.click()
            time.sleep(1)
            webdriver.ActionChains(driver).send_keys(city_search).perform()
            time.sleep(2)
            webdriver.ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ENTER).perform()
            print(f"  ✅ City: {city_full}")
        except Exception as e:
            print(f"  ⚠️ City hatası: {e}")
        time.sleep(2)

        try:
            date_inp = driver.find_element(By.XPATH, "//input[@placeholder='Date']")
            driver.execute_script("arguments[0].click();", date_inp)
            time.sleep(1)
            date_inp.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
            today = datetime.now().strftime("%d.%m.%Y")
            date_inp.send_keys(today, Keys.ENTER)
            print(f"  ✅ Tarih: {today}")
        except Exception as e:
            print(f"  ⚠️ Tarih hatası: {e}")
        time.sleep(2)

        try:
            job_ctrl = wait.until(EC.element_to_be_clickable((By.XPATH,
                "//label[contains(text(),'Job Title') or contains(text(),'Job title')]"
                "/following-sibling::div//div[contains(@class,'Select-control')]"
                " | (//div[contains(@class,'Select-control')])[3]"
            )))
            job_ctrl.click()
            time.sleep(1)
            webdriver.ActionChains(driver).send_keys("Picker", Keys.ENTER).perform()
            print("  ✅ Job Title: Picker")
        except Exception as e:
            print(f"  ⚠️ Job Title hatası: {e}")
        time.sleep(2)

        save_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Save contract')]")))
        driver.execute_script("arguments[0].click();", save_btn)
        print("  ✅ Sözleşme kaydedildi.")
        time.sleep(5)

    except Exception as e:
        print(f"  ❌ Sözleşme hatası: {e}")
        raise


def _rooster_assign_depot(driver, wait, mapping_dict, depot_code):
    print("  🏭 Depo atanıyor...")
    try:
        overview_tab = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//a[contains(@href,'/overview')] | //span[text()='Overview']/ancestor::a"
        )))
        driver.execute_script("arguments[0].click();", overview_tab)
        time.sleep(5)

        vendor_id = mapping_dict.get(depot_code, {}).get("vendor_id", "")
        if not vendor_id:
            print("  ⚠️ Vendor ID bulunamadı.")
            return

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

        inp = wait.until(EC.presence_of_element_located((By.XPATH,
            "//input[contains(@id,'MultipleSelectWrapper') or contains(@id,'starting_points')]"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", inp)
        driver.execute_script("arguments[0].click();", inp)
        time.sleep(1)

        webdriver.ActionChains(driver).move_to_element(inp).click().send_keys(vendor_id).perform()
        time.sleep(4)
        webdriver.ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ENTER).perform()
        time.sleep(2)

        save_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-test='save-starting-points']")))
        driver.execute_script("arguments[0].click();", save_btn)
        print("  ✅ Depo ataması tamamlandı.")
        time.sleep(2)

    except Exception as e:
        print(f"  ❌ Depo atama hatası: {e}")
        raise


# ==========================================
# 🟣 FAZ 2: ÇARŞI PORTAL
# ==========================================

def carsi_create_picker(driver, wait, row_data, store_name, rooster_id):
    full_name = f"{row_data['name']} {row_data['surname']}"
    print(f"  🛍️ Çarşı Portal: {full_name} işleniyor...")

    # --- TEMEL BİLGİ GİRİŞİ ---
    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)

        create_btn = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//button[contains(normalize-space(text()),'Create New Picker')]"
        )))
        driver.execute_script("arguments[0].click();", create_btn)
        time.sleep(2)

        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='firstName']"))).send_keys(full_name)
        driver.find_element(By.XPATH, "//input[@id='roosterEmployeeId']").send_keys(rooster_id)
        driver.find_element(By.XPATH, "//input[@id='email']").send_keys(row_data['email'])
        driver.find_element(By.XPATH, "//input[@id='password']").send_keys("12345678")

        pass_confirm_inp = driver.find_element(By.XPATH, "//input[@id='passwordConfirmation']")
        pass_confirm_inp.clear()
        pass_confirm_inp.send_keys("12345678")
        print("    ✅ Temel veriler ve şifreler girildi.")
        time.sleep(2)

    except Exception as e:
        print(f"    ❌ Temel bilgi hatası: {e}")
        return False

    # --- MAĞAZA SEÇİM BÖLÜMÜ ---
    if store_name:
        clean_store_name = store_name.strip()
        print(f"    → Mağaza seçiliyor: '{clean_store_name}'")

        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)

            # Store input = passwordConfirmation'dan sonra gelen ilk text input
            store_input = wait.until(EC.presence_of_element_located((By.XPATH,
                "//input[@id='passwordConfirmation']/following::input[@type='text'][1]"
            )))
            print("    ✅ Store input bulundu")

            # Tıkla ve mağaza adını yaz
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", store_input)
            time.sleep(0.5)
            webdriver.ActionChains(driver).move_to_element(store_input).click().perform()
            time.sleep(0.5)

            store_input.send_keys(Keys.CONTROL + "a")
            store_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.3)

            print("    ⌨️ Mağaza adı yazılıyor...")
            for char in clean_store_name:
                store_input.send_keys(char)
                time.sleep(0.07)

            print("    🔍 Liste bekleniyor...")
            time.sleep(3)

            # --- LİSTEDEN SEÇ ---
            selected = False

            search_keywords = [clean_store_name]
            if ',' in clean_store_name:
                search_keywords.append(clean_store_name.split(',')[0].strip())
                after = re.sub(r'\(.*?\)', '', clean_store_name.split(',')[1]).strip()
                if after:
                    search_keywords.append(after)
            m = re.search(r'\(([^)]+)\)', clean_store_name)
            if m:
                search_keywords.append(m.group(1))

            print(f"    🔑 Denenen keyword'ler: {search_keywords}")

            for kw in search_keywords:
                if selected or not kw or len(kw) < 3:
                    continue
                safe_kw = kw.replace("'", "\\'")

                try:
                    el = driver.execute_script(f"""
                        var kw = '{safe_kw}';

                        // 1. Açık dialog içinde ara (aria-hidden=false)
                        var dialogs = document.querySelectorAll('[role="dialog"]');
                        for (var d = 0; d < dialogs.length; d++) {{
                            if (dialogs[d].getAttribute('aria-hidden') !== 'false') continue;
                            var nodes = dialogs[d].querySelectorAll('*');
                            for (var i = 0; i < nodes.length; i++) {{
                                if (nodes[i].children.length === 0
                                    && nodes[i].textContent.trim().includes(kw)) {{
                                    return nodes[i];
                                }}
                            }}
                        }}

                        // 2. Tüm sayfada tam metin eşleşmesi olan görünür element
                        var all = document.querySelectorAll('span, li, div, button');
                        for (var j = 0; j < all.length; j++) {{
                            var rect = all[j].getBoundingClientRect();
                            if (rect.width === 0 || rect.height === 0) continue;
                            if (all[j].children.length === 0
                                && all[j].textContent.trim() === kw) {{
                                return all[j];
                            }}
                        }}

                        return null;
                    """)

                    if el:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        time.sleep(0.4)
                        driver.execute_script("arguments[0].click();", el)
                        time.sleep(0.5)
                        print(f"    ✅ Mağaza seçildi (kw: '{kw}')")
                        selected = True
                        break

                except Exception as je:
                    print(f"    ⚠️ JS hatası: {je}")

            # Son çare: ARROW_DOWN + ENTER
            if not selected:
                print("    ⚠️ JS ile bulunamadı, ARROW_DOWN deneniyor...")
                try:
                    webdriver.ActionChains(driver).move_to_element(store_input).click().perform()
                    time.sleep(0.3)
                    store_input.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.5)
                    store_input.send_keys(Keys.ENTER)
                    print("    ✅ ARROW_DOWN + ENTER ile seçildi.")
                    selected = True
                except Exception as e2:
                    print(f"    ❌ ARROW_DOWN da başarısız: {e2}")

            if not selected:
                print("    ❌ Mağaza seçilemedi.")
                return False

            # Seçim doğrulama
            time.sleep(1)
            try:
                confirm = driver.find_element(By.XPATH,
                    f"//span[contains(@class,'u26') and contains(.,'{clean_store_name[:15]}')]"
                )
                print(f"    ✅ Seçim doğrulandı: '{confirm.text}'")
            except Exception:
                print("    ⚠️ Seçim etiketi görülemedi, devam ediliyor...")

        except Exception as store_err:
            print(f"    ❌ Mağaza seçim hatası: {store_err}")
            traceback.print_exc()
            return False

    # --- CREATE BUTONUNA TIK ---
    try:
        create_final_btn = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//button[normalize-space(text())='Create' and contains(@class,'u73')]"
            " | //button[normalize-space(text())='Create']"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", create_final_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", create_final_btn)
        print("    ✅ Create butonuna tıklandı.")
        time.sleep(3)
        return True

    except Exception as e:
        print(f"    ❌ Create butonu tıklanamadı: {e}")
        return False
# ==========================================
# 🚀 ANA SÜREÇ
# ==========================================

def run():
    print("\n🚀 ROOSTER + ÇARŞI ROBOTU BAŞLATILIYOR...\n")

    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=opts)
    wait   = WebDriverWait(driver, 30)

    try:
        print("☁️ Google Sheets bağlanıyor...")
        client     = get_sheets_client()
        mapping_ws = client.open_by_url(MAP_SHEET_URL).worksheet(MAP_TAB_NAME)
        data_ws    = client.open_by_url(DATA_SHEET_URL).worksheet(DATA_TAB_NAME)

        mapping_dict = {
            tr_char_replace(r[0]): {
                "dmart_store": r[1].strip() if len(r) > 1 else "",
                "name":        r[2].strip() if len(r) > 2 else "",
                "city":        r[3].strip() if len(r) > 3 else "Istanbul",
                "vendor_id":   r[4].strip() if len(r) > 4 else "",
            }
            for r in mapping_ws.get_all_values()[1:]
            if r and r[0].strip()
        }
        print(f"  ✅ Mapping yüklendi: {len(mapping_dict)} depo.")

        # ==========================================
        # 🔵 FAZ 1: ROOSTER
        # ==========================================
        print("\n🔵 ========== FAZ 1: ROOSTER ==========")

        if not ensure_logged_in(driver, ROOSTER_URL):
            print("❌ Rooster girişi başarısız.")
            return

        dismiss_popups(driver, wait)

        rows = data_ws.get_all_values()

        for idx, row in enumerate(rows):
            row_num = idx + 1
            if row_num == 1:
                continue

            email_val  = safe_col(row, COL_EMAIL)
            status_val = safe_col(row, COL_STATUS)

            if "TAMAMLANDI" in status_val or not email_val:
                continue

            success, result = rooster_create_worker(driver, wait, row, mapping_dict)

            if success:
                if result:
                    data_ws.update_cell(row_num, COL_EMP_ID, result)
                    print(f"  ✅ Employee ID '{result}' → P{row_num} yazıldı.")
                data_ws.update_cell(row_num, COL_STATUS, "TAMAMLANDI")
                print(f"  ✅ Satır {row_num} Rooster TAMAMLANDI.")
            else:
                data_ws.update_cell(row_num, COL_STATUS, f"HATA: {result[:80]}")
                print(f"  ❌ Satır {row_num} Rooster HATA.")
                try:
                    driver.get(ROOSTER_URL)
                    time.sleep(5)
                except Exception:
                    pass

        # ==========================================
        # 🟣 FAZ 2: ÇARŞI PORTAL
        # ==========================================
        print("\n🟣 ========== FAZ 2: ÇARŞI PORTAL ==========")

        if not ensure_logged_in(driver, CARSI_URL):
            print("❌ Çarşı Portal girişi başarısız.")
            return

        print("🔍 Çarşı pop-up kapatılıyor (ESC)...")
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1.5)
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)

        rows = data_ws.get_all_values()

        for idx, row in enumerate(rows):
            row_num = idx + 1
            if row_num == 1:
                continue

            status_val    = safe_col(row, COL_STATUS)
            emp_id_val    = safe_col(row, COL_EMP_ID)
            portal_status = safe_col(row, COL_PORTAL_ST)

            if "TAMAMLANDI" not in status_val:
                continue
            if not emp_id_val:
                print(f"  ⚠️ Satır {row_num}: Employee ID yok, Portal atlandı.")
                continue
            if "PORTAL: OK" in portal_status:
                continue

            depot_code = tr_char_replace(safe_col(row, COL_DEPOT))
            store_name = mapping_dict.get(depot_code, {}).get("dmart_store", "")

            row_data = {
                "name":    safe_col(row, COL_NAME),
                "surname": safe_col(row, COL_SURNAME),
                "email":   safe_col(row, COL_EMAIL),
            }

            driver.get(CARSI_URL)
            time.sleep(4)

            success = carsi_create_picker(driver, wait, row_data, store_name, emp_id_val)

            if success:
                data_ws.update_cell(row_num, COL_PORTAL_ST, "PORTAL: OK")
                print(f"  ✅ Satır {row_num} Çarşı Portal TAMAMLANDI.")
            else:
                data_ws.update_cell(row_num, COL_PORTAL_ST, "HATA: Portal")
                print(f"  ❌ Satır {row_num} Çarşı Portal HATA.")

        print("\n🎉 TÜM SÜREÇ TAMAMLANDI!")

    except Exception as e:
        print(f"\n💥 GENEL HATA: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    run()
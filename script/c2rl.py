# ------------- IMPORTS -------------
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import configparser
import time
import sys
import re
import os
import threading
import pystray
from pystray import MenuItem as item
from PIL import Image
import logging
from logging.handlers import RotatingFileHandler
import math
import win32com.client
import random
import subprocess
import ctypes



# ------------- GLOBALS -------------
app_name = "c2rl"
app_description = "Codetwo License Reset"
app_version = "v1.3"
app_author = "https://github.com/mrtnRitter"

driver = None
target_url = None
user_data_dir = None
profile_dir = None
discover_timeout = None
watchdog_timeout = None
debug = None

base_path = None
idle_str = "Warte auf Daten ..."
timeout_seconds = 0
menu_timeout_str = idle_str
menu_license_str = idle_str
app_status = "default"



# ------------- RESOURCE PATH -------------
def resource_path(relative_path):
    """ 
    Get the absolute path to a resource, works for both frozen and non-frozen applications.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)



# ------------- FUNCTIONS -------------
def init():
    """
    Setup and configure the script.
    """
    global target_url
    global user_data_dir
    global profile_dir
    global discover_timeout
    global watchdog_timeout
    global debug
    global base_path

    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
        check_autostart_lnk()
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    setup_logging()
    logging.info(f"Starting {app_name} {app_version}")

    config_path = os.path.join(base_path, "config.ini")
    if not os.path.exists(config_path):
        make_config(config_path)
        exit()

    cfg = parse_config(config_path)
    if not cfg:
        make_config(config_path)
        cfg = parse_config(config_path)
        if not cfg:
            logging.critical("Config file could not be read after re-creation - App quit!")
            exit()
    
    target_url, user_data_dir, profile_dir, discover_timeout, watchdog_timeout, debug = cfg

    if not target_url:
        logging.error("Tenant ID is missing in the config file - App quit!")
        exit()

    if not user_data_dir or not profile_dir:
        add_browser_profile_to_config(config_path)


def check_autostart_lnk():
    """
    Create autostart entry for EXE if needed.
    """
    autostart = os.path.join(os.getenv('APPDATA'), r"Microsoft\Windows\Start Menu\Programs\Startup")
    exe_path = os.path.join(base_path, f"{app_name}.exe")
    shortcut = os.path.join(autostart, f"{app_name}.lnk")

    if not os.path.exists(shortcut):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut_obj = shell.CreateShortCut(shortcut)
        shortcut_obj.Targetpath = exe_path
        shortcut_obj.WorkingDirectory = base_path
        shortcut_obj.save()



def setup_logging():
    """
    Setup logging configuration.
    """
    log_path = os.path.join(base_path, "c2rl_log.log")
    logging.basicConfig(
        level=logging.INFO,
        format="[%(levelname)s] %(asctime)s %(threadName)s %(funcName)s: %(message)s",
        handlers=[RotatingFileHandler(
            log_path,
            maxBytes=10*1024*1024,
            backupCount=5,
            encoding="utf-8"
            )]
    )



def add_browser_profile_to_config(config_path):
    """
    Add the browser profile path to the config file.
    """
    logging.info("No browser profile in config file, attempting to add it.")
    
    global user_data_dir
    global profile_dir
    global driver
    
    setup_driver(headless=True)
    driver.get("chrome://version")
    profile_path = driver.find_element(By.ID, "profile_path").text
    driver.quit()
    driver = None

    user_data_dir = os.path.dirname(os.path.dirname(profile_path))
    profile_dir = os.path.basename(os.path.dirname(profile_path))

    config = configparser.ConfigParser()
    config.read(config_path)
    config["DEFAULT"]["user_data_dir"] = user_data_dir
    config["DEFAULT"]["profile_dir"] = profile_dir
    with open(config_path, "w") as f:
        config.write(f)



def make_config(config_path):
    """
    Create a default config file.
    """
    logging.info("No config file found, creating from scratch.")
    logging.info("Add tenant ID in config file to use the app.")

    with open(config_path, "w") as f:
        f.write(f"# Config file for {app_name} {app_version}\n")
        f.write("[DEFAULT]\n")
        f.write("tenant =\n")
        f.write("user_data_dir =\n")
        f.write("profile_dir =\n")
        f.write("discover_timeout = 10\n")
        f.write("watchdog_timeout = 600\n")
        f.write("debug = False\n")



def parse_config(config_path):
    """
    Parse the config file.
    """
    logging.info("Reading config file...")

    config = configparser.ConfigParser()
    config.read(config_path)

    try:
        tenant = config.get("DEFAULT", "tenant")
        if tenant:
            target_url = f"https://emailsignatures365.codetwo.com/dashboard/tenants/{tenant}/licenses"
        else:
            target_url = None
        user_data_dir = config.get("DEFAULT", "user_data_dir")
        profile_dir = config.get("DEFAULT", "profile_dir")
        discover_timeout = config.getint("DEFAULT", "discover_timeout")
        watchdog_timeout = config.getint("DEFAULT", "watchdog_timeout")
        debug = config.getboolean("DEFAULT", "debug")
        return target_url, user_data_dir, profile_dir, discover_timeout, watchdog_timeout, debug

    except Exception as e:
        logging.error(f"Error reading config file: {e}")
        return False



def setup_driver(headless):
    """
    Setup the Chrome driver with options.
    """
    global driver
    
    if driver:
        driver.quit()
        driver = None
    
    if not internet_available():
        return False
    
    logging.info("Setting up browser...")
    
    options = Options()

    if debug:
        headless = False
        options.add_experimental_option("detach", True)
    else:
        headless = headless
    
    if headless:
        options.add_argument("--headless")
    
    options.add_argument("--log-level=3")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--window-size=800,1000")

    if user_data_dir and profile_dir:
        options.add_argument(f"user-data-dir={user_data_dir}")
        options.add_argument(f"profile-directory={profile_dir}")

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(discover_timeout)

    driver.get(target_url)
    time.sleep(discover_timeout)
    logging.info("Browser ready.")
    return True



def auto_login():
    """
    Fullfil login procedure.
    """
    if not internet_available():
        return False
    
    logging.info("Attempting to auto login...")
    try:
        driver.find_element(By.ID, "AzureADCommonSigninExchange").click()
        driver.find_element(By.XPATH, "//small[text()='Angemeldet']").click()
        logging.info("Auto login successful.")
        return True
        
    except Exception as e:
        logging.error("Auto login failed: " + str(e).splitlines()[0])
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        driver.save_screenshot(os.path.join(base_path, f"auto_login_{timestamp}.png"))
        return False



def manual_login():
    """
    Open the browser for manual login and session renewal.
    """
    global driver
    if driver:
        driver.quit()
        driver = None
    
    if not internet_available():
        return False
    
    if setup_driver(headless=False):
        logging.info("Open Browser and waiting for user to manual login...")

        while True:
            try:
                driver.title
                time.sleep(discover_timeout)
            except WebDriverException:
                driver = None
                break



def get_menu_license_str(app):
    """
    Get the license count from the page.
    """
    logging.info("Attempting to fetch license count from page...")

    global menu_license_str
    lic_usage = "N/A"

    if not internet_available():
        result = False
    
    else:
        try:
            driver.refresh()
            time.sleep(discover_timeout)
            lic_usage = driver.find_element(By.TAG_NAME, "dd").text
            result = True
        
        except NoSuchElementException as e:
            logging.error("License counter not found: " + str(e).splitlines()[0])
            result = False

    menu_license_str = f"Lizenzen: {lic_usage}"
    app.menu = build_menu()
    logging.info(f"{menu_license_str} currenently in use.")
    return result



def reset_license_counter():
    """
    Reset the license count.
    """
    if not internet_available():
        return False

    logging.info("Attempting to click license reset button...")
    
    try:
        driver.refresh()
        time.sleep(discover_timeout)
        btns = driver.find_elements(By.CSS_SELECTOR, "button.c2-button.c2-button--solid")
        for btn in btns:
            if btn.find_element(By.TAG_NAME, "span").get_attribute("outerHTML") == '<span class="display-contents">Reset license count (signature)</span>':
                btn.click()
                time.sleep(5)
                logging.info("License reset button clicked.")
                return True
            
    except Exception as e:
        logging.error("License reset failed: " + str(e).splitlines()[0])
        return False
    
    logging.warning(f"License reset button not found on page.")
    return False



def get_timeout_seconds():
    """
    Get last reset date and return timeout to next reset. Always run reset_license_counter() first.
    """    
    logging.info("Attempting to fetch reset lock timeout...")
    
    global timeout_seconds

    if not internet_available():
        return False
    
    try:
        msg = driver.find_element(By.CLASS_NAME, "c2-message-text").text
        match = re.search(r"on ([A-Za-z]+\s\d{1,2},\s\d{4}) at ([\d:]+\s[AP]M\sUTC)", msg)
        if match:
                date_str = match.group(1)
                time_str = match.group(2)

                dt_str = f"{date_str} {time_str.replace(' UTC', '')}"
                utc_dt = datetime.strptime(dt_str, "%B %d, %Y %I:%M %p")
                utc_dt = utc_dt.replace(tzinfo=ZoneInfo("UTC"))
                berlin_dt = utc_dt.astimezone(ZoneInfo("Europe/Berlin"))

                last_reset_date = berlin_dt.strftime("%d.%m.%Y")
                last_reset_time = berlin_dt.strftime("%H:%M")
                timeout_seconds = int(math.ceil(((berlin_dt + timedelta(hours=24)) - datetime.now(ZoneInfo("Europe/Berlin"))).total_seconds())) + random.randint(60, 180)
                driver.find_element(By.CLASS_NAME, "c2-dlg-button ").click()

                logging.info(f"Last reset was: {last_reset_date}, at: {last_reset_time}, reset locked for: {timeout_seconds} seconds.")
                time.sleep(5)
                return True

        else:
                logging.error(f"Could not find any dates in: '{msg}'")
                return False    

    except NoSuchElementException as e:
        logging.error("Reset lock timeout not found: " + str(e).splitlines()[0])
        return False



def update_reset_lock_timeout(app):
    """
    Set the reset lock timeout in the UI.
    """
    logging.info("Updating reset lock timeout in UI...")
    
    global timeout_seconds
    global menu_timeout_str

    while timeout_seconds > 0:
        h, m = timeout_seconds // 3600, (timeout_seconds % 3600) // 60

        if m >= 30:
            h += 1

        if timeout_seconds >= 3600:
            menu_timeout_str = f"Reset in {int(h)}h"
            update_interval = 3600
        elif timeout_seconds >= 600:
            menu_timeout_str = f"Reset in {int(m)}m"
            update_interval = 600
        elif timeout_seconds >= 60:
            menu_timeout_str = f"Reset in {int(m)}m"
            update_interval = 60

        app.menu = build_menu()

        sleep_time = min(update_interval, timeout_seconds)  
        logging.info(f"Reset locked for: {timeout_seconds} seconds, next update in {sleep_time} seconds.")

        time_before_sleep = int(time.time())
        time.sleep(sleep_time)
        time_after_sleep = int(time.time())
        timeout_seconds -= sleep_time + (sleep_time - (time_after_sleep - time_before_sleep))

        

def timeout_and_reset(app):
    """
    Handles license reset, reset lock timeout and updates the UI.
    """
    while True:
        if setup_driver(headless=True):
            if reset_license_counter():
                if get_timeout_seconds():
                    update_reset_lock_timeout(app)

            elif auto_login():
                time.sleep(discover_timeout)
                continue
                    
            else:
                manual_login()

        time.sleep(discover_timeout)



def license_watchdog(app):
    """
    Watchdog thread to monitor the license counter.
    """
    global app_status

    while True:
        while timeout_seconds > 60:
            if setup_driver(headless=True):
                if get_menu_license_str(app):
                    sleep_time = min(watchdog_timeout, timeout_seconds)
                    time.sleep(sleep_time)

                elif auto_login():
                    time.sleep(discover_timeout)
                    continue
            
                else:
                    manual_login()
            
            time.sleep(discover_timeout)
        time.sleep(discover_timeout)



def internet_available():   
    """
    Check if the internet connection is available.
    """
    global driver
    global app_status
    global app

    if ctypes.windll.wininet.InternetGetConnectedState(0, 0) == 0:
        logging.warning("Internet connection is not available.")
        
        if driver:
            driver.quit()
            driver = None

        if app_status == "default":      
            set_tray_icon(app, "error")
            app_status = "error"

        return False
    
    else:
        if app_status == "error":
            set_tray_icon(app, "default")
            app_status = "default"

        return True



# -------- UI FUNCTIONS -------------

def set_tray_icon(app, status):
    if status == "default":
        app.icon = Image.open(ico_default)
    elif status == "error":
        app.icon = Image.open(ico_error)

def get_settings():
    if base_path:
        os.startfile(base_path)

def get_about():
    os.startfile("https://github.com/mrtnRitter/c2rl")

def on_quit(app):
    global driver
    if driver:
        driver.quit()
        driver = None
    logging.info("App closed by user.")
    app.stop()

def build_menu():
    return pystray.Menu(
        item(f"{menu_license_str}", None),
        item(f"{menu_timeout_str}", None),
        pystray.Menu.SEPARATOR,
        item("Einstellungen", get_settings),
        item("Über", get_about),
        item("Beenden", on_quit)
    )



# ------------- MAIN -------------
if __name__ == "__main__":

    ico_default = resource_path("res/ico_default.ico")
    ico_error = resource_path("res/ico_error.ico")

    init()
    app = pystray.Icon(app_description, Image.open(ico_default), app_description + " " + app_version, build_menu())

    timeout_and_reset_t = threading.Thread(target=timeout_and_reset, args=(app,), daemon=True)
    timeout_and_reset_t.start()

    license_watchdog_t = threading.Thread(target=license_watchdog, args=(app,), daemon=True)
    license_watchdog_t.start()

    app.run()
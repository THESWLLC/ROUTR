import time
import threading
import logging
import os
import sys
import pandas as pd  
import re
import json
import base64
import math
from datetime import datetime, timedelta
import usaddress

import customtkinter as ctk
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox

from webdriver_manager.chrome import ChromeDriverManager
import psutil  # For process cleanup
import concurrent.futures

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    NoSuchElementException,
    ElementClickInterceptedException,
    NoSuchWindowException
)
from selenium.webdriver.support.expected_conditions import staleness_of
from urllib3.exceptions import ProtocolError
from http.client import RemoteDisconnected
from http.server import HTTPServer, BaseHTTPRequestHandler
import requests

# â”€â”€â”€ OPS Package Tracking Exception List API â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
OPS_API_BASE = os.getenv("ROUTR_API_BASE", "https://api.example.com")
OPS_UI_BASE = os.getenv("ROUTR_UI_BASE", "https://app.example.com")


def get_ops_exception_list(station, date_str, track_type, exception_code, carrierx_id, bearer_token, timeout=30):
    """
    Call the OPS package-tracking-exception-list API. Returns list of rows (ptelData).
    station: e.g. 'hwoa'
    date_str: YYYY-MM-DD
    track_type: e.g. 'STAT'
    exception_code: e.g. '13'
    carrierx_id: employee ID for employee-id query param
    bearer_token: SSO Bearer token (no 'Bearer ' prefix required).
    """
    url = (
        f"{OPS_API_BASE}/api/v3/package-tracking-exception-list/"
        f"{station}/{date_str}/{track_type}/{exception_code}"
    )
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Authorization": f"Bearer {bearer_token}" if bearer_token and not bearer_token.startswith("Bearer ") else bearer_token or "",
        "Referer": f"{OPS_UI_BASE}/",
    }
    params = {"employee-id": carrierx_id}
    resp = requests.get(url, headers=headers, params=params, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return data.get("ptelData", [])


# â”€â”€â”€ IDs that always get access â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
BYPASS_ACCESS_IDS = set()  # Removed private allowlist for public showcase

# â”€â”€â”€ Hard-coded Google key & helper â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")  # Set in local env only

# â”€â”€â”€ Custom Tooltip Class â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class CustomTooltip:
    """A custom tooltip class that works with customtkinter widgets."""
    
    def __init__(self, widget, text="", delay=500, wraplength=300):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.wraplength = wraplength
        self.tooltip_window = None
        self.id_after = None
        
        # Bind events
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.widget.bind("<Motion>", self.on_motion)
    
    def on_enter(self, event=None):
        """Called when mouse enters the widget."""
        if self.id_after:
            self.widget.after_cancel(self.id_after)
        self.id_after = self.widget.after(self.delay, self.show_tooltip)
    
    def on_leave(self, event=None):
        """Called when mouse leaves the widget."""
        if self.id_after:
            self.widget.after_cancel(self.id_after)
            self.id_after = None
        self.hide_tooltip()
    
    def on_motion(self, event=None):
        """Called when mouse moves within the widget."""
        if self.tooltip_window:
            self.update_tooltip_position(event)
    
    def show_tooltip(self):
        """Show the tooltip window."""
        if self.tooltip_window or not self.text:
            return
        
        # Get mouse position and place tooltip slightly ABOVE the widget
        # so it doesn't appear "under" the widget window (especially in popouts).
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() - 30
        
        # Create tooltip window
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        # Ensure tooltip appears ABOVE any parent toplevels (incl. widget pop-out)
        try:
            # Make tooltip a transient child of the top-level window
            top = self.widget.winfo_toplevel()
            self.tooltip_window.transient(top)
        except Exception:
            pass
        try:
            # Force it on top of the window stack
            self.tooltip_window.lift()
            self.tooltip_window.attributes("-topmost", True)
        except Exception:
            pass
        
        # Configure tooltip appearance
        label = tk.Label(
            self.tooltip_window,
            text=self.text,
            justify='left',
            background='#333333',
            foreground='white',
            relief='solid',
            borderwidth=1,
            font=('Arial', 9),
            wraplength=self.wraplength,
            padx=8,
            pady=6
        )
        label.pack()
    
    def hide_tooltip(self):
        """Hide the tooltip window."""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def update_tooltip_position(self, event):
        """Update tooltip position based on mouse movement."""
        if self.tooltip_window:
            # Keep tooltip horizontally following the mouse, but fixed ABOVE
            # the widget instead of snapping below on movement.
            x = self.widget.winfo_rootx() + event.x + 20
            y = self.widget.winfo_rooty() - 30
            self.tooltip_window.wm_geometry(f"+{x}+{y}")
    
    def update_text(self, new_text):
        """Update the tooltip text."""
        self.text = new_text

def normalize_with_geocode(raw_addr: str, zip_code: str) -> str:
    """
    Hits Google's Geocoding API with the raw street + ZIP,
    returns the first formatted_address.
    """
    params = {
        "address": raw_addr,
        "components": f"postal_code:{zip_code}",
        "key": GOOGLE_API_KEY,
    }
    resp = requests.get(
        "https://maps.googleapis.com/maps/api/geocode/json",
        params=params,
        timeout=5
    ).json()
    if resp.get("status") != "OK" or not resp.get("results"):
        raise ValueError(f"Geocode failed: {resp.get('status')}")
    return resp["results"][0]["formatted_address"]

# â”€â”€â”€ Animated Loading Spinner Counter â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class AnimatedSpinnerCounter:
    """Counter that shows either an animated spinner or text value."""
    
    def __init__(self, parent, size=30, color="#FF6600", font_size=None):  # Carrier orange
        self.parent = parent
        self.size = size
        self.color = color
        self.angle = 0
        self.is_spinning = False
        self.current_text = ""
        
        # Use provided font_size or calculate from spinner size
        if font_size is None:
            font_size = int(size * 0.67)  # Default ratio
        
        # Create frame to hold both spinner and text
        self.frame = tk.Frame(parent, bg="black")
        
        # Create canvas for the spinner
        self.canvas = tk.Canvas(
            self.frame, 
            width=size, 
            height=size, 
            bg="black", 
            highlightthickness=0
        )
        
        # Create text label (initially hidden)
        self.text_label = tk.Label(
            self.frame,
            text="",
            font=("Tahoma", font_size, "bold"),
            fg=color,
            bg="black"
        )
        
        # Draw the spinner bars
        self._draw_spinner()
        
        # Start with blank text (no numbers until process starts)
        self.show_text("")
        
    def _draw_spinner(self):
        """Draw the spinner with circular arrangement of bars like the image."""
        # Clear canvas
        self.canvas.delete("all")
        
        # Calculate center and radius
        center_x = self.size / 2
        center_y = self.size / 2
        radius = (self.size - 4) / 2
        
        # Number of bars (16 like in the image)
        num_bars = 16
        bar_width = 2
        bar_length = 6
        
        # Draw each bar
        for i in range(num_bars):
            # Calculate angle for this bar
            angle = (i * 360 / num_bars) + self.angle
            
            # Convert to radians
            rad = math.radians(angle)
            
            # Calculate bar start and end points
            start_x = center_x + (radius - bar_length) * math.cos(rad)
            start_y = center_y + (radius - bar_length) * math.sin(rad)
            end_x = center_x + radius * math.cos(rad)
            end_y = center_y + radius * math.sin(rad)
            
            # Determine bar color (most bars are light grey, 2 adjacent are darker)
            if i == int(self.angle / (360 / num_bars)) % num_bars or i == (int(self.angle / (360 / num_bars)) + 1) % num_bars:
                bar_color = "#666666"  # Darker grey for the 2 adjacent bars
            else:
                bar_color = "#CCCCCC"  # Light grey for other bars
            
            # Draw the bar
            self.canvas.create_line(
                start_x, start_y, end_x, end_y,
                fill=bar_color,
                width=bar_width,
                capstyle=tk.ROUND
            )
    
    def start_spinner(self):
        """Start the spinning animation."""
        if not self.is_spinning:
            self.is_spinning = True
            self._animate()
    
    def stop_spinner(self):
        """Stop the spinning animation."""
        self.is_spinning = False
    
    def _animate(self):
        """Animate the spinner rotation."""
        if self.is_spinning:
            self.angle = (self.angle + 10) % 360  # Rotate 10 degrees
            self._draw_spinner()
            self.parent.after(50, self._animate)  # Update every 50ms
    
    def show_spinner(self):
        """Show the spinner and hide text."""
        self.text_label.pack_forget()
        self.canvas.pack()
        self.start_spinner()
    
    def show_text(self, text):
        """Show text and hide spinner."""
        self.stop_spinner()
        self.canvas.pack_forget()
        self.current_text = str(text)
        self.text_label.configure(text=self.current_text)
        self.text_label.pack()
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)
    
    def pack_forget(self):
        """Unpack the frame."""
        self.frame.pack_forget()
    
    def configure(self, **kwargs):
        """Configure the counter (for compatibility with CTkLabel)."""
        if "text" in kwargs:
            # If text is provided, show the text
            self.show_text(kwargs["text"])
        return self.frame.configure(**kwargs)

# â”€â”€â”€ Simple Tooltip for hover popups â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.showtip()

    def leave(self, event=None):
        self.hidetip()

    def showtip(self):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, justify=tk.LEFT,
            background="#ffffe0", relief=tk.SOLID, borderwidth=1,
            font=("tahoma", "8", "normal")
        )
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        if tw:
            tw.destroy()
        self.tipwindow = None


# â”€â”€â”€ Logging & Constants â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
log_file_path = os.path.join(os.path.expanduser("~"), "ROUTR_autoclear.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode='a', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

BRAND_PURPLE           = "#4D148C"
BRAND_ORANGE           = "#FF6600"
ROUTR_EXTENSION_PORT   = 51999
IDLE_STATUS_TEXT       = "Monitoring for new messages..."
CLEAR_IDLE_STATUS_TEXT = "Monitoring for new Clear messages..."
POLL_FREQ              = 5  # seconds between idle polls

# XPaths & selectors for Auto-Clear
NEW_MSG_XPATH = "//td//span[contains(normalize-space(.),'WORKAREA') and contains(normalize-space(.),'REQUEST TO')]"
MSG_TAB_XPATH      = "//span[starts-with(normalize-space(.), 'Messages - ')]"
INBOX_TAB_XPATH    = "//span[starts-with(normalize-space(.), 'INBOX')]"
REFRESH_BTN_CSS    = "span.p-button-icon.pi.pi-refresh"
SUBMIT_BTN_XPATH   = "//span[normalize-space(text())='Submit']"
SYS_MSGS_TAB_XPATH = "//span[starts-with(normalize-space(.),'SYSTEM MSGS')]"
CREATE_MSG_TAB_XPATH = "//span[normalize-space(text())='CREATE MESSAGE']"


# â”€â”€â”€ Auto-Read Monitor â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class CarrierSeleniumMonitor:
    def _set_agent_display_name_from_full(self, full_name: str):
        """
        Given a EmployeeDirectory full name like "Shawn Williams",
        store a compact display form such as "Shawn W." for CXPC signatures.
        """
        try:
            name = full_name.replace('\u00A0', ' ').strip()
            if not name:
                return
            parts = name.split()
            if len(parts) == 1:
                display = parts[0].title()
            else:
                first = parts[0].title()
                last_initial = parts[-1][0].upper()
                display = f"{first} {last_initial}."
            self.agent_display_name = display
            logging.info(f"âœ… Dispatcher display name set to {self.agent_display_name!r}")
            # Update the Auto-Clear signature label on the Tk main thread
            try:
                if self.clear_signature_label is not None:
                    def _update():
                        # Mirror into manual entry if empty so the user sees current value
                        try:
                            if self.clear_name_entry is not None and not self.clear_name_entry.get().strip():
                                self.clear_name_entry.delete(0, "end")
                                self.clear_name_entry.insert(0, self.agent_display_name)
                        except Exception:
                            pass
                        if self.agent_display_name:
                            txt = f"Sending clear message as -CXPC {self.agent_display_name}"
                        else:
                            txt = "Sending clear message without CXPC name (login to EmployeeDirectory to populate)."
                        self.clear_signature_label.configure(text=txt)
                    # Ensure we run on the UI thread
                    self.after(0, _update)
            except Exception as e:
                logging.warning(f"âš ï¸ Failed to update clear-signature label: {e!r}")
        except Exception as e:
            logging.warning(f"âš ï¸ Failed to parse dispatcher display name from {full_name!r}: {e!r}")

    def _format_signed_message(self, text: str) -> str:
        """
        Append the CXPC signature to any outbound courier message generated
        by the Autoâ€‘Read monitor, unless it's already present or we don't
        yet know the dispatcher name.
        """
        base = (text or "").rstrip()
        name = getattr(self, "agent_display_name", "").strip()
        if not name:
            return base
        signature = f"-CXPC {name}"
        if base.endswith(signature) or signature in base:
            return base
        sep = "" if base.endswith((".", "!", "?")) else "."
        return f"{base}{sep} {signature}"

    def check_access(self, username: str) -> bool:
        """
        Returns True if the given username appears in EmployeeDirectory
        with region "Dispatch / Southern Region", else False.
        """
        # â”€â”€â”€ Bypass GUI access check for special IDs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if username.strip() in BYPASS_ACCESS_IDS:
            return True

        # install() returns the path to the chromedriver binary
        driver_path = ChromeDriverManager().install()
        # send chromedriver logs to nul (on Windows) / devnull
        service     = Service(driver_path, log_path=os.devnull)
        opts    = Options()
        opts.add_argument("--disable-gpu")
        opts.add_argument("--headless")
        opts.add_argument("--window-size=1920,1080")
        # kill all chromium logging below FATAL
        opts.add_argument("--log-level=3")
        # drop the "enable-logging" switch entirely
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver  = webdriver.Chrome(service=service, options=opts)
        try:
            driver.get("https://internal.example.com")
            wait = WebDriverWait(driver, 10)
            # find the search box, enter username, submit
            search = wait.until(
                EC.element_to_be_clickable((By.NAME, "Search"))
            )
            search.clear()
            search.send_keys(username, Keys.RETURN)
            # wait for the results to load
            # (you may need to tweak this selector if the page structure changes)
            wait.until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//font[@face='Arial' and normalize-space(text())='Dispatch / Southern Region']"
                ))
            )
            return True
        except TimeoutException:
            return False
        finally:
            driver.quit()
            
    def __init__(self, username, password, dispatch_location, stop_event,
                 message_types, reply_mode=False, early_pu_reply=False,
                 auto_walkup=False,
                 update_status=None, finish_callback=None,
                 log_callback=None,
                 mfa_factor_callback=None, mfa_code_callback=None,
                 agent_display_name: str | None = None):
        self.username          = username
        self.password          = password
        self.dispatch_location = dispatch_location.upper()
        self.stop_event        = stop_event
        self.message_types     = message_types
        self.update_status     = update_status
        self.reply_mode        = reply_mode
        self.early_pu_reply    = early_pu_reply
        self.auto_walkup       = auto_walkup
        self.finish_callback   = finish_callback
        self.log_callback      = log_callback
        self.mfa_factor_callback = mfa_factor_callback
        self.mfa_code_callback   = mfa_code_callback
        # Dispatcher display name used for signing Autoâ€‘Read replies
        self.agent_display_name = agent_display_name or ""
        self.message_actions = {
            "Pickup Reminder": {
                "xpath": "//td[.//span[contains(normalize-space(.), 'Pickup Reminder')]]"
            },
            "Early PU": {
                "xpath": "//td[.//span[contains(normalize-space(.), 'Early PU')]]"
            },
            "No Pickup List": {
                "xpath": "//td[.//span[contains(normalize-space(.), 'No Pickup List')]]"
            },
        }
        self.driver = None
        self._walkup_tab_open = False

    @staticmethod
    def ensure_home_loaded(driver, page_timeout: int = 10, ui_timeout: int = 3):
        """
        1) Wait up to `page_timeout` seconds for document.readyState to be 'complete'.
        2) Then wait up to `ui_timeout` seconds for the FRO menu to be clickable.
        """
        # 1) full load
        WebDriverWait(driver, page_timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        # 2) first meaningful UI element
        WebDriverWait(driver, ui_timeout).until(
            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='FRO']"))
        ) 

    # â”€â”€â”€ Dynamic input finder helper â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    @staticmethod
    def find_username_input(driver, wait):
        """Dynamically find the username input field by checking multiple strategies."""
        try:
            # Strategy 1: Look for inputs with username-related attributes (most reliable)
            # Use XPath for case-insensitive matching
            username_xpaths = [
                "//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'user')]",
                "//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'login')]",
                "//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'email')]",
                "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'user')]",
                "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'login')]",
                "//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'user')]",
                "//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'email')]",
                "//input[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'user')]",
                "//input[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'email')]"
            ]
            for xpath in username_xpaths:
                try:
                    elements = driver.find_elements(By.XPATH, xpath)
                    for elem in elements:
                        if elem.is_displayed() and elem.get_attribute("type") != "password":
                            return elem
                except:
                    continue
            
            # Also try CSS selectors (case-sensitive)
            username_selectors = [
                "input[name*='user']",
                "input[name*='login']",
                "input[name*='email']",
                "input[id*='user']",
                "input[id*='login']",
                "input[placeholder*='user']",
                "input[placeholder*='email']",
                "input[aria-label*='user']",
                "input[aria-label*='email']"
            ]
            for selector in username_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if elem.is_displayed() and elem.get_attribute("type") != "password":
                            return elem
                except:
                    continue
            
            # Strategy 2: Find password field first, then look for text input before it
            password_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
            if password_inputs:
                # Find all text/email inputs on the page
                text_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[type='email']")
                # Get the one that appears before the password field by checking DOM order
                for text_input in text_inputs:
                    if text_input.is_displayed():
                        try:
                            # Check if this input appears before password in DOM
                            is_before = driver.execute_script("""
                                var password = arguments[0];
                                var textInput = arguments[1];
                                var position = password.compareDocumentPosition(textInput);
                                return (position & Node.DOCUMENT_POSITION_PRECEDING) !== 0;
                            """, password_inputs[0], text_input)
                            if is_before:
                                return text_input
                        except:
                            # Fallback: check Y position
                            try:
                                text_y = text_input.location['y']
                                pass_y = password_inputs[0].location['y']
                                if text_y < pass_y:
                                    return text_input
                            except:
                                pass
                # If no input before password, return the first visible text input
                for text_input in text_inputs:
                    if text_input.is_displayed():
                        return text_input
            
            # Strategy 3: Find first visible text/email input
            text_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[type='email']")
            for inp in text_inputs:
                if inp.is_displayed():
                    return inp
            
            # Strategy 4: Fallback to any input that's not password
            all_inputs = driver.find_elements(By.CSS_SELECTOR, "input")
            for inp in all_inputs:
                inp_type = inp.get_attribute("type") or ""
                if inp_type != "password" and inp.is_displayed():
                    return inp
                    
        except Exception as e:
            logging.warning(f"âš ï¸ Error finding username input: {e}")
        
        return None
    
    @staticmethod
    def find_password_input(driver, wait):
        """Dynamically find the password input field."""
        try:
            # Most reliable: find by type="password"
            password_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
            for inp in password_inputs:
                if inp.is_displayed():
                    return inp
            
            # Fallback: look for password-related attributes (case-insensitive with XPath)
            password_xpaths = [
                "//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pass')]",
                "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pass')]",
                "//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pass')]",
                "//input[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pass')]"
            ]
            for xpath in password_xpaths:
                try:
                    elements = driver.find_elements(By.XPATH, xpath)
                    for elem in elements:
                        if elem.is_displayed():
                            return elem
                except:
                    continue
            
            # Also try CSS selectors (case-sensitive)
            password_selectors = [
                "input[name*='pass']",
                "input[id*='pass']",
                "input[placeholder*='pass']",
                "input[aria-label*='pass']"
            ]
            for selector in password_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if elem.is_displayed():
                            return elem
                except:
                    continue
                    
        except Exception as e:
            logging.warning(f"âš ï¸ Error finding password input: {e}")
        
        return None
    
    # â”€â”€â”€ Helper function to check if already logged in â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    @staticmethod
    def check_if_logged_in(driver):
        """Check if the user is already logged in by checking URL and FRO menu."""
        try:
            # Check if URL contains /home
            if "/home" in driver.current_url:
                # Verify FRO menu is present to confirm full login
                try:
                    fro_menu = driver.find_element(By.XPATH, "//span[normalize-space(text())='FRO']")
                    if fro_menu.is_displayed():
                        return True
                except:
                    pass
            return False
        except:
            return False

    # â”€â”€â”€ Custom ExpectedCondition for dynamic username field detection â”€â”€â”€â”€â”€â”€â”€
    class UsernameFieldFound:
        """Custom ExpectedCondition that waits for username field to appear using dynamic detection."""
        def __call__(self, driver):
            try:
                element = CarrierSeleniumMonitor.find_username_input(driver, None)
                if element and element.is_displayed():
                    return element
            except (StaleElementReferenceException, NoSuchElementException):
                pass
            return False
    
    class PasswordFieldFound:
        """Custom ExpectedCondition that waits for password field to appear using dynamic detection."""
        def __call__(self, driver):
            try:
                element = CarrierSeleniumMonitor.find_password_input(driver, None)
                if element and element.is_displayed():
                    return element
            except (StaleElementReferenceException, NoSuchElementException):
                pass
            return False

    # â”€â”€â”€ SSO / Identity Portal login helper â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def sso_navigate(self, driver, wait,
                      url="https://internal.example.com",
                      username=None, password=None):
        driver.get(url)

        # â€” Poll frequently for auto-login (automatic login via SSO/cookies can take ~10 seconds) â€”
        logging.info("ðŸ” Checking for auto-login (polling every 1 second for up to 15 seconds)...")
        for attempt in range(15):  # Check for up to 15 seconds
            if CarrierSeleniumMonitor.check_if_logged_in(driver):
                logging.info(f"âœ… Auto-login detected after {attempt + 1} seconds (URL contains /home and FRO menu found)")
                return  # Already logged in, skip login process
            time.sleep(1)
        
        # Final check after polling period
        if CarrierSeleniumMonitor.check_if_logged_in(driver):
            logging.info("âœ… Auto-login detected after polling period")
            return  # Already logged in, skip login process

        # â€” only attempt login if SSO form appears within 30 s â€”
        try:
            short_wait = WebDriverWait(driver, 30)
            
            # Wait for username input field to appear using dynamic detection (with frequent auto-login checks)
            logging.info("ðŸ” Waiting for username input field to appear...")
            usr = None
            start_time = time.time()
            timeout = 30
            check_interval = 1  # Check every 1 second
            
            while time.time() - start_time < timeout:
                # Check for auto-login first
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected during username field wait")
                    return  # Already logged in, skip login process
                
                # Try to find username field
                try:
                    usr = CarrierSeleniumMonitor.find_username_input(driver, None)
                    if usr and usr.is_displayed():
                        break
                except:
                    pass
                
                time.sleep(check_interval)
            
            if not usr:
                # Before raising error, final check for auto-login
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected (final check)")
                    return  # Already logged in, skip login process
                
                # Fallback: try original method
                try:
                    usr = short_wait.until(CarrierSeleniumMonitor.UsernameFieldFound())
                except TimeoutException:
                    # Final check for auto-login before raising exception
                    if CarrierSeleniumMonitor.check_if_logged_in(driver):
                        logging.info("âœ… Auto-login detected during username wait timeout")
                        return  # Already logged in, skip login process
                    
                    # Fallback to waiting for any clickable input (original behavior)
                    logging.warning("âš ï¸ Could not dynamically find username input, trying fallback...")
                    try:
                        # Try common input ID patterns
                        for input_id in ["input44", "input28", "input"]:
                            try:
                                usr = short_wait.until(EC.element_to_be_clickable((By.ID, input_id)))
                                break
                            except:
                                continue
                        # If still not found, get first visible text input
                        if not usr:
                            usr = short_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
                    except:
                        # Final check for auto-login before raising exception
                        if CarrierSeleniumMonitor.check_if_logged_in(driver):
                            logging.info("âœ… Auto-login detected (fallback check)")
                            return  # Already logged in, skip login process
                        raise Exception("Could not find username input field")
            
            logging.info(f"âœ… Found username input: {usr.get_attribute('id') or usr.get_attribute('name') or 'unknown'}")
            usr.clear()
            usr.send_keys(username)
            
            # Look for Next button - try multiple possible selectors
            next_button = None
            next_selectors = [
                "input[type='submit']",
                "button[type='submit']", 
                "input[value*='Next']",
                "button:contains('Next')",
                ".next-button",
                "#next-button",
                "input[value='Next']",
                "button[value='Next']"
            ]
            
            for selector in next_selectors:
                try:
                    if selector.startswith("button:contains"):
                        # Use XPath for text content
                        next_button = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]")))
                    else:
                        next_button = short_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    break
                except TimeoutException:
                    continue
            
            if next_button:
                logging.info("ðŸ”„ Found Next button, clicking...")
                next_button.click()
                # Wait a moment for the page to transition
                time.sleep(2)
            else:
                logging.info("ðŸ”„ No Next button found, trying Enter key...")
                usr.send_keys(Keys.RETURN)
                time.sleep(2)

            # Wait for password input field to appear using dynamic detection (with frequent auto-login checks)
            logging.info("ðŸ” Waiting for password input field to appear...")
            pw = None
            start_time = time.time()
            timeout = 30
            check_interval = 1  # Check every 1 second
            
            while time.time() - start_time < timeout:
                # Check for auto-login first
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected during password field wait")
                    return  # Already logged in, skip login process
                
                # Try to find password field
                try:
                    pw = CarrierSeleniumMonitor.find_password_input(driver, None)
                    if pw and pw.is_displayed():
                        break
                except:
                    pass
                
                time.sleep(check_interval)
            
            if not pw:
                # Before raising error, final check for auto-login
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected (final check)")
                    return  # Already logged in, skip login process
                
                # Fallback: try original method
                try:
                    pw = short_wait.until(CarrierSeleniumMonitor.PasswordFieldFound())
                except TimeoutException:
                    # Final check for auto-login before raising exception
                    if CarrierSeleniumMonitor.check_if_logged_in(driver):
                        logging.info("âœ… Auto-login detected during password wait timeout")
                        return  # Already logged in, skip login process
                    
                    # Fallback to waiting for password input (original behavior)
                    logging.warning("âš ï¸ Could not dynamically find password input, trying fallback...")
                    try:
                        # Try common input ID patterns
                        for input_id in ["input70", "input54", "input"]:
                            try:
                                pw = short_wait.until(EC.element_to_be_clickable((By.ID, input_id)))
                                break
                            except:
                                continue
                        # If still not found, get password type input
                        if not pw:
                            pw = short_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
                    except:
                        # Final check for auto-login before raising exception
                        if CarrierSeleniumMonitor.check_if_logged_in(driver):
                            logging.info("âœ… Auto-login detected (fallback check)")
                            return  # Already logged in, skip login process
                        raise Exception("Could not find password input field")
            
            logging.info(f"âœ… Found password input: {pw.get_attribute('id') or pw.get_attribute('name') or 'unknown'}")
            pw.clear()
            pw.send_keys(password + Keys.RETURN)

            # â”€â”€â”€ Existing MFAâ€factor selection and code/push flow â”€â”€
            factors = short_wait.until(
                EC.presence_of_all_elements_located((
                    By.CSS_SELECTOR, "a.button.select-factor.link-button"
                ))
            )
            factor = self.mfa_factor_callback()

            # click the link they chose
            for a in factors:
                lbl = a.get_attribute("aria-label").lower()
                if "enter a code" in lbl and factor == "code":
                    a.click(); break
                if "push notification" in lbl and factor == "push":
                    a.click(); break

            # if they chose the code-entry flow, ask for the TOTP and submit it
            if factor == "code":
                # 1) prompt your CTk dialog for the TOTP code
                code = self.mfa_code_callback()

                # 2) wait for the TOTP input to appear (use name="credentials.totp")
                code_in = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.NAME, "credentials.totp"))
                )
                code_in.clear()
                code_in.send_keys(code)

                # 3) click the Verify/Submit button if it exists
                try:
                    verify_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((
                            By.CSS_SELECTOR,
                            'button[type="submit"], input[type="submit"]'
                        ))
                    )
                    verify_btn.click()
                except TimeoutException:
                    # fallback: press Enter in the input field
                    code_in.send_keys(Keys.RETURN)

        except TimeoutException:
            # no MFA step appeared â†’ continue
            pass

        # â”€â”€â”€ Wait up to 60 s for SSO to finish MFA and land on /home â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        try:
            # 1) Wait for the URL to contain "/home" (covers push approval)
            WebDriverWait(driver, 60).until(EC.url_contains("/home"))
            # 2) Then wait for the FRO menu to be clickable (UI is fully loaded)
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='FRO']"))
            )
        except TimeoutException:
            raise TimeoutException("SSO never navigated to home after MFA")

    def run(self):
        opts = Options()
        opts.add_argument("--disable-extensions")
        opts.add_argument("--headless")  # Enabled headless mode for auto-read process
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--disable-gpu")
        
        # Bot detection evasion techniques
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option('useAutomationExtension', False)
        
        # Add realistic user agent
        opts.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

        # â”€â”€â”€ Helper to click tabs reliably â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        def click_tab(xpath, name):
            # 1) wait for any blocking overlay to disappear
            try:
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-widget-overlay"))
                )
            except:
                pass
            # 2) wait for the tab element to exist
            tab = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            # 3) scroll into view & attempt click (JS fallback)
            driver.execute_script("arguments[0].scrollIntoView(true);", tab)
            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                tab.click()
            except:
                driver.execute_script("arguments[0].click();", tab)
            logging.info(f"ðŸ”€ Switched to {name} tab")

        # retry loop: keep going until stopped
        while not self.stop_event.is_set():
            driver = None
            try:
                logging.info(f"ðŸ” Starting Auto-Read monitor for {self.dispatch_location}")
                driver_path = ChromeDriverManager().install()
                service     = Service(driver_path, log_path=os.devnull)
                self.driver = webdriver.Chrome(service=service, options=opts)
                driver = self.driver
                wait   = WebDriverWait(driver, 30)

                # â”€â”€â”€ LOGIN (attempt only if form appears) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                logging.info("âž¡ï¸  Navigating to pickup UI")
                self.sso_navigate(
                    driver, wait,
                    "https://internal.example.com",
                    self.username, self.password
                )

                # â”€â”€â”€ FAST home-page detection with up to 5 retries â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                for attempt in range(5):
                    # 1) poll URL for '/home' up to ~2â€‰s
                    for _ in range(20):
                        if "/home" in driver.current_url:
                            break
                        time.sleep(0.1)

                    # 2) try waiting up to 1â€‰s for the FRO menu
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='FRO']"))
                        )
                        logging.info(f"âœ… Home UI detected on attempt {attempt+1}")
                        break  # success!
                    except TimeoutException:
                        logging.warning(f"â— Home detection attempt {attempt+1} failed")
                        if attempt < 4:
                            logging.info("ðŸ”„ Refreshing page and retryingâ€¦")
                            driver.refresh()
                            time.sleep(7)
                        else:
                            # all retries exhaustedâ€”let outer handler catch it
                            raise

                logging.info("âœ… Login successful (home UI detected)")

                # Remove automation indicators to avoid bot detection
                try:
                    driver.execute_script("""
                        // Remove webdriver property
                        Object.defineProperty(navigator, 'webdriver', {
                            get: () => undefined,
                        });
                        
                        // Remove automation properties
                        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
                        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                        
                        // Override permissions
                        const originalQuery = window.navigator.permissions.query;
                        window.navigator.permissions.query = (parameters) => (
                            parameters.name === 'notifications' ?
                                Promise.resolve({ state: Notification.permission }) :
                                originalQuery(parameters)
                        );
                    """)
                    logging.info("âœ… Removed automation indicators")
                except Exception as e:
                    logging.info(f"â„¹ï¸ Could not remove automation indicators: {e}")

                if self.update_status:
                    self.update_status(self.dispatch_location, "Login successful")

                # â”€â”€â”€ NAVIGATE TO MESSAGE LIST â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Message List']"))).click()
                # poll for the Message-List URL (max ~5 s)
                for _ in range(50):
                    if "/home/pickup/(messages:messages)" in driver.current_url:
                        break
                    time.sleep(0.1)
                fld = wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                fld.clear()
                fld.send_keys(self.dispatch_location + Keys.RETURN)
                if self.dispatch_location == "BCTA":
                    time.sleep(3)
                logging.info(f"âœ… Dispatch location set to {self.dispatch_location}")
                if self.update_status:
                    self.update_status(self.dispatch_location, f"Location set: {self.dispatch_location}")

                # kick the UI into "idle" immediately,
                # so user sees "Monitoring for new messagesâ€¦" before first poll
                if self.update_status:
                    self.update_status(self.dispatch_location, IDLE_STATUS_TEXT)

                actions = ActionChains(driver)

                # â”€â”€â”€ MONITORING LOOP â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                while not self.stop_event.is_set():
                    action_taken = False
                    
                    # â”€â”€â”€ Check SYS MSGS tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    try:
                        click_tab(SYS_MSGS_TAB_XPATH, "SYS MSGS")
                        # allow UI to settle before scanning for messages
                        time.sleep(0.5)
                    except Exception as e:
                        logging.warning(f"âš ï¸ Could not switch to SYS MSGS tab: {e}")
                        # Try refreshing the page and retrying, but never kill the monitor
                        try:
                            driver.refresh()
                            time.sleep(3)
                            # Wait for page to load and try again
                            wait.until(
                                EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")
                                )
                            )
                        except Exception as refresh_error:
                            logging.error(f"âŒ Failed to refresh page (ignoring): {refresh_error}")
                        # Either way, skip this iteration and retry the loop
                        continue

                    # â”€â”€â”€ Auto-WalkUp detection â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    logging.info(f"ðŸ” Auto-WalkUp enabled: {self.auto_walkup}")
                    if self.auto_walkup:
                        try:
                            # 0) refresh the list so new WalkUps show up immediately
                            try:
                                driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                # wait for any overlay to clear
                                WebDriverWait(driver, 3).until(
                                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-widget-overlay"))
                                )
                            except Exception:
                                pass
                            # 1) locate non-empty "#wu" spans and snapshot their text
                            xpath = (
                                "//span["
                                "contains(translate(normalize-space(.),"
                                " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                                " 'abcdefghijklmnopqrstuvwxyz'), '#wu')"
                                " or contains(translate(normalize-space(.),"
                                " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                                " 'abcdefghijklmnopqrstuvwxyz'), '#walkup')"
                                " or contains(translate(normalize-space(.),"
                                " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                                " 'abcdefghijklmnopqrstuvwxyz'), '#walk up')"
                                "]"
                            )
                            try:
                                raw_spans = driver.find_elements(By.XPATH, xpath)
                                logging.info(f"ðŸ” Found {len(raw_spans)} potential WalkUp spans")
                            except Exception as e:
                                logging.warning(f"âš ï¸ Failed to find WalkUp spans: {e}")
                                raw_spans = []
                            
                            # Pair each element with its initial text, so we can fall back if it goes stale
                            walkup_data = []
                            for s in raw_spans:
                                try:
                                    if s.text.strip():
                                        # Get the text first before any operations
                                        preview_text = s.text.strip()
                                        
                                        # Get route information BEFORE opening the message (to avoid stale elements)
                                        try:
                                            row = s.find_element(By.XPATH, "./ancestor::tr")
                                            cells = row.find_elements(By.TAG_NAME, "td")
                                            
                                            # Find the "From" column index
                                            header_ths = driver.find_elements(By.CSS_SELECTOR, "table thead tr th")
                                            from_idx = next(
                                                (i for i, th in enumerate(header_ths)
                                                if "from" in th.text.strip().lower()),
                                                None
                                            )
                                            if from_idx is not None:
                                                raw_route = cells[from_idx].text.strip()
                                                m_route = re.search(r"\d{1,4}", raw_route)
                                                route = m_route.group(0) if m_route else None
                                                logging.info(f"ðŸ” Found route {route} for walkup message")
                                            else:
                                                route = None
                                                logging.warning("âš ï¸ Could not find 'From' column header")
                                        except Exception as e:
                                            route = None
                                            logging.warning(f"âš ï¸ Could not get route from row: {e}")
                                        
                                        try:
                                            # Double-click the message row to open it and get full text
                                            row = s.find_element(By.XPATH, "./ancestor::tr")
                                            ActionChains(driver).double_click(row).perform()
                                            time.sleep(1.0)  # Wait for message to open
                                            
                                            try:
                                                # Look for the td element with the full message text
                                                full_text_td = driver.find_element(By.CSS_SELECTOR, "td[style*='white-space: pre']")
                                                if full_text_td.is_displayed():
                                                    full_msg = full_text_td.text.strip()
                                                    logging.info(f"ðŸ” Got full text from opened message: {full_msg}")
                                                else:
                                                    full_msg = preview_text
                                                    logging.info(f"ðŸ” Full text td not visible, using preview: {full_msg}")
                                            except:
                                                full_msg = preview_text
                                                logging.info(f"ðŸ” No full text td found, using preview: {full_msg}")
                                            
                                            # Process the walkup immediately while the message is still open
                                            walkup_data.append((s, full_msg, route))
                                            logging.info("ðŸ” Added walkup data for processing")
                                            
                                            # Click CLOSE button to return to messages
                                            try:
                                                close_btn = driver.find_element(By.XPATH, "//span[@class='p-button-label' and text()='CLOSE']")
                                                close_btn.click()
                                                time.sleep(0.5)  # Wait for messages to reload
                                                logging.info("ðŸ” Clicked CLOSE button to return to messages")
                                            except Exception as e:
                                                logging.warning(f"âš ï¸ Could not click CLOSE button: {e}")
                                        except Exception as e:
                                            logging.warning(f"âš ï¸ Could not get full message for span: {e}")
                                            walkup_data.append((s, preview_text, route))
                                except Exception as e:
                                    logging.warning(f"âš ï¸ Failed to get span text: {e}")
                                    continue
                        except Exception as e:
                            logging.error(f"âŒ Auto-WalkUp section failed: {type(e).__name__}: {str(e)}")
                            # Try to recover by refreshing the page
                            try:
                                driver.refresh()
                                time.sleep(3)
                                # Wait for page to load and try again
                                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                            except Exception as refresh_error:
                                logging.error(f"âŒ Failed to refresh page after Auto-WalkUp section error: {refresh_error}")
                            
                            logging.info(f"ðŸ” WalkUp data: {[text for _, text, _ in walkup_data]}")

                            # 2) find the FROM column index once
                            header_ths = driver.find_elements(By.CSS_SELECTOR, "table thead tr th")
                            from_idx = next((i for i, th in enumerate(header_ths)
                                            if th.text.strip().lower() == "from"), None)
                            if from_idx is None:
                                logging.error("ðŸš¨ Couldn't find FROM column â€“ skipping WalkUp.")
                            else:
                                for span, full_msg, route in walkup_data:
                                    # We already have the full message and route from the double-click extraction
                                    # No need to hover or re-locate elements
                                    
                                    if not route:
                                        logging.warning("âš ï¸ No route found for walkup message, skipping.")
                                        continue

                                # 5) ZIP (5 digits)
                                m_zip    = re.search(r"\b(\d{5})\b", full_msg)
                                zip_code = m_zip.group(1) if m_zip else ""
                            
                                # missing/invalid ZIP â†’ ask courier to resubmit, then archive & skip
                                if not re.fullmatch(r"\d{5}", zip_code):
                                    logging.warning(f"âš ï¸ Missing/Bad ZIP ({zip_code!r}) â€“ requesting resubmit")
                                    row = span.find_element(By.XPATH, "./ancestor::tr")
                                    try:
                                        ActionChains(driver).double_click(row).perform()
                                    except:
                                        driver.execute_script("arguments[0].click();", row)
                                        ta = WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, 'textarea[formcontrolname="newMessageContent"]'))
                                        )
                                        ta.clear()
                                        reply_body = (
                                            "Zip code missing or invalid. Please ensure your ZIP code is correct "
                                            "and resubmit your Walk Up request. Thanks."
                                        )
                                        reply_body = self._format_signed_message(reply_body)
                                        ta.send_keys(reply_body)
                                    send_btn = WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[normalize-space(text())='SEND']]"))
                                    )
                                    send_btn.click()
                                    # archive original request
                                    hist_icon = row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                    driver.execute_script("arguments[0].click();", hist_icon)
                                    # return to Messages tab
                                    driver.find_element(By.XPATH, INBOX_TAB_XPATH).click()
                                    time.sleep(0.5)
                                    driver.find_element(By.XPATH, MSG_TAB_XPATH).click()
                                    continue
                            
                                # strip off "#wu" and trailing ZIP to get raw street portion
                                raw_addr = re.sub(r"#wu", "", full_msg, flags=re.IGNORECASE).strip()
                                # 6a) Extract ZIP only if it appears at end-of-string
                                m_zip_end = re.search(r"(\b\d{5}\b)\s*$", full_msg)
                                if m_zip_end:
                                    raw_addr = raw_addr[: m_zip_end.start()].strip()
                            
                                # 6b) Normalize address with geocoding from the start, then parse
                                address = None
                                try:
                                    # Use geocoding as primary method for consistent normalization
                                    normalized = normalize_with_geocode(raw_addr, zip_code)
                                    logging.info(f"âœ… Geocoded address: {normalized}")

                                    # Parse the normalized address
                                    tagged, _ = usaddress.tag(normalized)
                                    if tagged.get("AddressNumber") and tagged.get("StreetName"):
                                        num = tagged.get("AddressNumber", "")
                                        pre_dir = tagged.get("StreetNamePreDirectional", "")
                                        street = tagged.get("StreetName", "")
                                        suffix = tagged.get("StreetNamePostType", "")
                                        post_dir = tagged.get("StreetNamePostDirectional", "")
                                        address = " ".join(
                                            p for p in (num, pre_dir, street, suffix, post_dir) if p
                                        ).strip()
                                        logging.info(f"âœ… Standardized address: {address}")
                                    else:
                                        logging.warning("âš ï¸ Can't parse geocoded address: defaulting to 'Walk Up'")
                                        address = "Walk Up"
                                except Exception as e:
                                    logging.warning(
                                        f"âš ï¸ Geocoding failed: {e!r}; defaulting to 'Walk Up'"
                                    )
                                    address = "Walk Up"

                                try:
                                    # 7) dismiss overlays
                                    WebDriverWait(driver, 5).until(
                                        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-widget-overlay"))
                                    )

                                    # 8) Pickup Request â†’ Quick Pickup
                                    if not self._walkup_tab_open:
                                        # first WalkUp: open via the menu
                                        menu = driver.find_element(
                                            By.XPATH,
                                            "//span[@class='ui-menuitem-text' and normalize-space(text())='Pickup Request']"
                                        )
                                        try:
                                            menu.click()
                                        except ElementClickInterceptedException:
                                            driver.execute_script("arguments[0].click();", menu)
                                        self._walkup_tab_open = True
                                    else:
                                        # subsequent WalkUps: focus Pickup Request by its text (fallback to menu if needed)
                                        try:
                                            driver.find_element(
                                                By.XPATH,
                                                "//span[normalize-space(text())='Pickup Request']"
                                            ).click()
                                        except NoSuchElementException:
                                            # if the tab header isn't found, re-open via the menu
                                            menu = driver.find_element(
                                                By.XPATH,
                                                "//span[@class='ui-menuitem-text' and normalize-space(text())='Pickup Request']"
                                            )
                                            try:
                                                menu.click()
                                            except ElementClickInterceptedException:
                                                driver.execute_script("arguments[0].click();", menu)

                                    # 1) Click "Validate Services" (scroll & JS-fallback)
                                    validate_btn = WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located((
                                            By.XPATH,
                                            "//button[contains(@class,'pickupButtons') and normalize-space(text())='Validate Services']"
                                        ))
                                    )
                                    # ensure it's in view
                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", validate_btn)
                                    try:
                                        validate_btn.click()
                                    except ElementClickInterceptedException:
                                        # overlay or animation blocking; use a JS click fallback
                                        driver.execute_script("arguments[0].click();", validate_btn)

                                        # 2) Enter ZIP into postalSvcValid
                                        svc_in = WebDriverWait(driver, 10).until(
                                            EC.element_to_be_clickable((
                                                By.CSS_SELECTOR,
                                                "input[formcontrolname='postalSvcValid']"
                                            ))
                                        )
                                        svc_in.clear()
                                        svc_in.send_keys(zip_code)

                                        # 3) Tab twice + ENTER to trigger validation
                                        for _ in range(2):
                                            ActionChains(driver).send_keys(Keys.TAB).perform()
                                            time.sleep(0.2)
                                        ActionChains(driver).send_keys(Keys.ENTER).perform()

                                        # 4) Wait for the Service Validation title
                                        WebDriverWait(driver, 10).until(
                                            EC.visibility_of_element_located((
                                                By.XPATH,
                                                "//span[contains(normalize-space(.),'Service Validation')]"
                                            ))
                                        )
                                        # Wait until locSvcValid has a value
                                        def non_empty_value(drv):
                                            inp = drv.find_element(
                                                By.CSS_SELECTOR,
                                                "input[formcontrolname='locSvcValid']"
                                            )
                                            val = inp.get_attribute("value").strip()
                                            return val or False

                                        validated_loc = WebDriverWait(driver, 10).until(non_empty_value).upper()
                                        logging.info(
                                            f"ðŸ” Service-validated location (from input): {validated_loc}"
                                        )

                                        # â”€â”€â”€ Validation mismatch handling â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                        if validated_loc != self.dispatch_location.upper():
                                            logging.warning(
                                                f"âš ï¸ Validation mismatch: auto-read='{self.dispatch_location.upper()}', "
                                                f"service-validate='{validated_loc}' â†’ sending mismatch reply"
                                            )
                                           # 1) Close the Service Validation pop-up
                                            try:
                                                close_btn = WebDriverWait(driver, 5).until(
                                                    EC.element_to_be_clickable((
                                                        By.XPATH,
                                                        "//button[contains(@class,'pickupButtons') and normalize-space(text())='Close']"
                                                    ))
                                                )
                                                # 2) scroll it into view and click it
                                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", close_btn)
                                                try:
                                                    close_btn.click()
                                                except ElementClickInterceptedException:
                                                    # fallback to JS click if normal click is blocked
                                                    driver.execute_script("arguments[0].click();", close_btn)
                                                logging.info("âœ… Service Validation pop-up closed via Close button")
                                            except TimeoutException:
                                                logging.warning("âš ï¸ Close button didn't appear in time; pop-up may already be gone")
                                            except Exception as e:
                                                logging.warning(f"âš ï¸ Error clicking Close button: {e!r}")
                                            except:
                                                pass

                                            # â”€â”€â”€ Close any Details tabs that might have been opened â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                            try:
                                                # Look for tabs with "Details" in the title and close them
                                                detail_tabs = driver.find_elements(
                                                    By.XPATH,
                                                    "//span[contains(@class, 'ui-tabview-title') and contains(text(), 'Details')]"
                                                )
                                                for detail_tab in detail_tabs:
                                                    try:
                                                        # Find the close button for this tab
                                                        close_btn = detail_tab.find_element(By.XPATH, "./following-sibling::span[contains(@class, 'ui-tabview-close')]")
                                                        close_btn.click()
                                                        logging.info("ðŸ”’ Closed Details tab")
                                                        time.sleep(0.5)
                                                    except Exception as close_error:
                                                        logging.warning(f"âš ï¸ Could not close Details tab: {close_error}")
                                            except Exception as e:
                                                logging.warning(f"âš ï¸ Could not find Details tabs to close: {e!r}")
                
                                            # â”€â”€â”€ Return to Messages tab and refresh the list with fallback â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                            for attempt in range(1, 4):
                                                try:
                                                    # focus Messages tab
                                                    driver.find_element(By.XPATH, MSG_TAB_XPATH).click()
                                                    time.sleep(1)
                                                    # click the in-UI Refresh button
                                                    driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                                    logging.info(f"ðŸ”„ Messages tab refreshed (attempt {attempt})")
                                                    break
                                                except Exception as e:
                                                    logging.warning(f"âš ï¸ Refresh attempt {attempt}/3 failed: {e!r}")
                                                    if attempt == 3:
                                                        # fallback: re-enter the dispatch-location filter
                                                        try:
                                                            fld = driver.find_element(
                                                                By.CSS_SELECTOR,
                                                                "input[formcontrolname='dispatchLocation']"
                                                            )
                                                            fld.click()
                                                            fld.clear()
                                                            fld.send_keys(self.dispatch_location + Keys.RETURN)
                                                            logging.info("ðŸ”„ Fallback refresh via dispatchLocation Enter")
                                                        except Exception as e2:
                                                            logging.error(f"âŒ Fallback refresh failed: {e2!r}")

                                            # 3) Re-find & open the original WalkUp row
                                            try:
                                                msg_row = driver.find_element(
                                                    By.XPATH,
                                                    f"//span[contains(normalize-space(.), \"{preview_text}\")]/ancestor::tr"
                                                )
                                                driver.execute_script(
                                                    "arguments[0].scrollIntoView({block:'center'});",
                                                    msg_row
                                                )
                                                ActionChains(driver).double_click(msg_row).perform()
                                            except Exception as e:
                                                logging.warning(f"âš ï¸ Could not open WalkUp message row: {e!r}")
                                                # skip reply if we can't open message
                                                continue

                                            # 4) Now type & send the mismatch reply
                                            try:
                                                ta = WebDriverWait(driver, 10).until(
                                                    EC.element_to_be_clickable((
                                                        By.CSS_SELECTOR,
                                                        'textarea[formcontrolname="newMessageContent"]'
                                                    ))
                                                )
                                                ta.clear()
                                                reply_body = (
                                                    "Zip code provided is for a different station. "
                                                    "Please ensure you are using the correct zip code within the station. "
                                                    "If you think this is an error send a new message. Thanks"
                                                )
                                                reply_body = self._format_signed_message(reply_body)
                                                ta.send_keys(reply_body)
                                                send_btn = WebDriverWait(driver, 10).until(
                                                    EC.element_to_be_clickable((
                                                        By.XPATH,
                                                        "//button[.//span[normalize-space(text())='SEND']]"
                                                    ))
                                                )
                                                send_btn.click()
                                                logging.info("âœ… WalkUp mismatch reply sent")
                                                if self.log_callback:
                                                    self.log_callback(
                                                        f"{self.dispatch_location}: WalkUp mismatch reply for route {route}"
                                                    )
                                                # â”€â”€â”€ Return to INBOX â†’ Messages so we resume monitoring â”€â”€â”€
                                                driver.find_element(By.XPATH, INBOX_TAB_XPATH).click()
                                                time.sleep(0.5)
                                                driver.find_element(By.XPATH, MSG_TAB_XPATH).click()

                                            except Exception as e:
                                                logging.warning(f"âš ï¸ Failed to send mismatch reply: {e!r}")

                                            # 3) Skip the rest of this WalkUp flow
                                            continue

                                        # --- If validation passed, close the pop-up and proceed ---
                                        try:
                                            close_btn = WebDriverWait(driver, 5).until(
                                                EC.element_to_be_clickable((
                                                    By.XPATH,
                                                    "//button[contains(@class,'pickupButtons') and normalize-space(text())='Close']"
                                                ))
                                            )
                                            close_btn.click()
                                            logging.info("âœ… Service Validation pop-up closed via Close button")
                                        except:
                                            # popup may already be gone
                                            pass

                                    except Exception as e:
                                        logging.warning(
                                            f"âš ï¸ Service-validate step failed: {e!r}, skipping WalkUp."
                                        )
                                        # ensure the pop-up is closed
                                        try:
                                            driver.find_elements(
                                                By.CSS_SELECTOR,
                                                "span.ui-tabview-close.pi.pi-times"
                                            )[-1].click()
                                        except:
                                            pass
                                        continue

                                    # â”€â”€â”€ Quick Pickup button â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                    quick = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((
                                            By.XPATH,
                                            "//button[contains(@class,'pickupButtons') "
                                            "and normalize-space(text())='Quick Pickup']"
                                        ))
                                    )
                                    # â”€â”€â”€ ensure no modal overlay is blocking â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                    try:
                                        WebDriverWait(driver, 5).until(
                                            EC.invisibility_of_element_located(
                                                (By.CSS_SELECTOR, "div.ui-widget-overlay")
                                            )
                                        )
                                    except TimeoutException:
                                        # still try clicking anyway
                                        pass
                                    # scroll into view & attempt click with JS fallback
                                    driver.execute_script("arguments[0].scrollIntoView(true);", quick)
                                    try:
                                        quick.click()
                                    except ElementClickInterceptedException:
                                        driver.execute_script("arguments[0].click();", quick)

                                    # 9) locate inputs
                                    route_in = WebDriverWait(driver, 10).until(
                                        lambda d: d.find_element(By.CSS_SELECTOR, "input[formcontrolname='walkupRoute']")
                                    )
                                    addr_in  = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='walkupAddress']")
                                    zip_in   = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='walkupPostal']")

                                    # 10) fill & validate
                                    route_in.clear(); route_in.send_keys(route)
                                    zip_in.clear();   zip_in.send_keys(zip_code)
                                    addr_in.clear();  addr_in.send_keys(address)

                                    v0, v1, v2 = (
                                        route_in.get_attribute("value"),
                                        zip_in.get_attribute("value"),
                                        addr_in.get_attribute("value")
                                    )
                                    logging.info(f"ðŸ“ Form â†’ route={v0!r}, zip={v1!r}, address={v2!r}")
                                    if not (re.fullmatch(r"\d{1,4}", v0) and re.fullmatch(r"\d{5}", v1) and v2):
                                        logging.error("ðŸš¨ Form validation failed, skipping Submit")
                                        continue

                                    # â”€â”€â”€ Keypress Tab at page level until Submit is focused, then ENTER â”€â”€â”€
                                    submit_btn = driver.find_element(By.CSS_SELECTOR, "button.pickupButtons.right")
                                
                                    # Use ActionChains.send_keys so Tab is sent to the document, not just the input
                                    for _ in range(10):
                                        ActionChains(driver).send_keys(Keys.TAB).perform()
                                        time.sleep(0.2)
                                        # check if Submit is now the active element
                                        if driver.execute_script(
                                            "return document.activeElement === arguments[0];", submit_btn
                                        ):
                                            break
                                
                                    # Once focused, hit ENTER (again at page level)
                                    ActionChains(driver).send_keys(Keys.ENTER).perform()

                                    logging.info(f"âœ… Auto-WalkUp created for route {route}")
                                    if self.log_callback:
                                        self.log_callback(f"{self.dispatch_location}: WalkUp created for {route}")
                                    
                                    # 1) wait for the "created successfully" toast and log its location & stop ID
                                    toast = WebDriverWait(driver, 10).until(
                                        EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ui-toast-detail"))
                                    )
                                    toast_msg = toast.text.strip()
                                    # e.g. "Dispatch HWOA 3408 06-20-25 (FR) has been created successfully."
                                    m = re.search(r"Dispatch\s+(\w+)\s+(\d+)", toast_msg)
                                    if m:
                                        loc_code, stop_id = m.group(1), m.group(2)
                                        log_txt = f"âœ… Successfully created WalkUp ({loc_code} {stop_id})"
                                        logging.info(log_txt)
                                        if self.log_callback:
                                            self.log_callback(f"{self.dispatch_location}: {log_txt}")
                                    else:
                                        logging.warning(f"âš ï¸ Unrecognized WalkUp toast: {toast_msg}")
                        
                                    # 2) small pause so user can see toast
                                    time.sleep(5)
                            
                                    # 2.5) archive the WalkUp request row into history, retry up to 5Ã—
                                    archived = False
                                    for attempt in range(1, 6):
                                        try:
                                            # re-find the row by its preview_text (or other unique selector)
                                            row = driver.find_element(
                                                By.XPATH,
                                                f"//span[contains(normalize-space(.), \"{preview_text}\")]/ancestor::tr"
                                            )
                                            hist_icon = row.find_element(
                                                By.CSS_SELECTOR,
                                                "i.action-icon[title='Move to history']"
                                            )
                                            driver.execute_script("arguments[0].click();", hist_icon)

                                            # wait until that row disappears to be sure it actually archived
                                            WebDriverWait(driver, 5).until(staleness_of(row))

                                            logging.info(f"âœ… Archived WalkUp message for route {route} on attempt {attempt}")
                                            archived = True
                                            break

                                        except (StaleElementReferenceException, NoSuchElementException, TimeoutException) as e:
                                            logging.warning(f"âš ï¸ Archive WalkUp attempt {attempt}/5 failed: {e!r}")
                                            # refresh the message list to get the latest DOM
                                            try:
                                                driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                                time.sleep(1)
                                            except:
                                                pass

                                    if not archived:
                                        logging.error(f"âŒ Could not archive WalkUp message after 5 attempts for route {route}")
                            
                                    # â”€â”€â”€ Close any Details tabs that might have been opened â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                    try:
                                        # Look for tabs with "Details" in the title and close them
                                        detail_tabs = driver.find_elements(
                                            By.XPATH,
                                            "//span[contains(@class, 'ui-tabview-title') and contains(text(), 'Details')]"
                                        )
                                        for detail_tab in detail_tabs:
                                            try:
                                                # Find the close button for this tab
                                                close_btn = detail_tab.find_element(By.XPATH, "./following-sibling::span[contains(@class, 'ui-tabview-close')]")
                                                close_btn.click()
                                                logging.info("ðŸ”’ Closed Details tab")
                                                time.sleep(0.5)
                                            except Exception as close_error:
                                                logging.warning(f"âš ï¸ Could not close Details tab: {close_error}")
                                    except Exception as e:
                                        logging.warning(f"âš ï¸ Could not find Details tabs to close: {e!r}")
        
                                    # â”€â”€â”€ Return to Messages tab and refresh the list with fallback â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                    for attempt in range(1, 4):
                                        try:
                                            # focus Messages tab
                                            driver.find_element(By.XPATH, MSG_TAB_XPATH).click()
                                            time.sleep(1)
                                            # click the in-UI Refresh button
                                            driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                            logging.info(f"ðŸ”„ Messages tab refreshed (attempt {attempt})")
                                            break
                                        except Exception as e:
                                            logging.warning(f"âš ï¸ Refresh attempt {attempt}/3 failed: {e!r}")
                                            if attempt == 3:
                                                # fallback: re-enter the dispatch-location filter
                                                try:
                                                    fld = driver.find_element(
                                                        By.CSS_SELECTOR,
                                                        "input[formcontrolname='dispatchLocation']"
                                                    )
                                                    fld.click()
                                                    fld.clear()
                                                    fld.send_keys(self.dispatch_location + Keys.RETURN)
                                                    logging.info("ðŸ”„ Fallback refresh via dispatchLocation Enter")
                                                except Exception as e2:
                                                    logging.error(f"âŒ Fallback refresh failed: {e2!r}")

                                except Exception as e:
                                    logging.error(f"âŒ Auto-WalkUp flow failed: {type(e).__name__}: {str(e)}")
                                    # Try to recover by refreshing the page
                                    try:
                                        driver.refresh()
                                        time.sleep(3)
                                        # Wait for page to load and try again
                                        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                                    except Exception as refresh_error:
                                        logging.error(f"âŒ Failed to refresh page after WalkUp error: {refresh_error}")
                                # after handling any spans, continue on to the rest of your loop
                        except Exception as e:
                            logging.error(f"âŒ Auto-WalkUp section failed: {type(e).__name__}: {str(e)}")
                            # Try to recover by refreshing the page
                            try:
                                driver.refresh()
                                time.sleep(3)
                                # Wait for page to load and try again
                                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                            except Exception as refresh_error:
                                logging.error(f"âŒ Failed to refresh page after Auto-WalkUp section error: {refresh_error}")

                    action_taken = False

                    for mtype in self.message_types:
                        info = self.message_actions[mtype]

                        # â”€â”€â”€ pagination lookup â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        # â”€â”€â”€ pagination lookup â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        el = None
                        # first, ensure you're on page 1
                        try:
                            driver.find_element(By.CSS_SELECTOR, "a.ui-paginator-first").click()
                            time.sleep(0.3)
                        except:
                            pass

                        for _ in range(3):  # try pages 1â†’2â†’3
                            try:
                                el = WebDriverWait(driver, 2).until(
                                    EC.element_to_be_clickable((By.XPATH, info["xpath"]))
                                )
                                break
                            except (TimeoutException, StaleElementReferenceException):
                                # click "next page" if you can
                                try:
                                    nxt = driver.find_element(By.CSS_SELECTOR, "a.ui-paginator-next")
                                    nxt.click()
                                    time.sleep(0.5)
                                except:
                                    break
                        if not el:
                            continue
                        # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

                        # â”€â”€â”€ scroll into view and log â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        driver.execute_script("arguments[0].scrollIntoView(true);", el)
                        logging.info(f"ðŸ“¬ Found {mtype} â†’ {el.text}")

                        # â”€â”€â”€ handle Pickup Reminder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        if mtype == "Pickup Reminder":
                            # Extract route number before processing
                            route = None
                            try:
                                cells = el.find_elements(By.TAG_NAME, "td")
                                header_ths = driver.find_elements(By.CSS_SELECTOR, "table thead tr th")
                                from_idx = next(
                                    (i for i, th in enumerate(header_ths)
                                    if "from" in th.text.strip().lower()),
                                    None
                                )
                                if from_idx is not None and from_idx < len(cells):
                                    raw_route = cells[from_idx].text.strip()
                                    m_route = re.search(r"\d{1,4}", raw_route)
                                    route = m_route.group(0) if m_route else None
                            except Exception as e:
                                logging.warning(f"âš ï¸ Could not extract route for Pickup Reminder: {e}")
                            
                            if self.reply_mode:
                                # â”€â”€ RETRY OPENING THE REPLY PANE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                subj = None
                                for attempt in range(1, 4):
                                    try:
                                        # bring the row fully into view
                                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                            
                                        # try double-click, fallback to click/JS click
                                        try:
                                            ActionChains(driver).double_click(el).perform()
                                        except (ElementClickInterceptedException, StaleElementReferenceException):
                                            try:
                                                el.click()
                                            except Exception:
                                                driver.execute_script("arguments[0].click();", el)
                            
                                        # wait up to 5s for the subject field
                                        subj = WebDriverWait(driver, 5).until(
                                            EC.visibility_of_element_located((
                                                By.CSS_SELECTOR, "input[formcontrolname='newMessageSubject']"
                                            ))
                                        )
                                        logging.info(f"âœ… Reply pane opened on attempt {attempt}")
                                        break
                                    except (TimeoutException, StaleElementReferenceException) as e:
                                        logging.warning(f"âš ï¸ Attempt {attempt}/3: reply pane not ready ({e!r})")
                                        time.sleep(1)
                                if not subj:
                                    logging.error("âŒ Could not open Pickup Reminder reply pane after 3 attempts â€” archiving")
                                    # give the UI a quick refresh before archiving, to avoid stale state
                                    try:
                                        driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                        time.sleep(1)
                                    except:
                                        pass
                                    try:
                                        row = el.find_element(By.XPATH, "./ancestor::tr")
                                        hist_icon = WebDriverWait(row, 5).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                "i.action-icon[title='Move to history']"
                                            ))
                                        )
                                        driver.execute_script("arguments[0].click();", hist_icon)
                                        if self.update_status:
                                            self.update_status(self.dispatch_location, "Pickup Reminder moved to history")
                                    except Exception as e:
                                        logging.error(f"âš ï¸ Failed to archive after pane-open failure: {e!r}")
                                    continue  # skip to the next message

                                # â”€â”€ NOW YOUR ORIGINAL REPLY LOGIC â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

                                # clear any old subject text
                                try:
                                    subj.click()
                                    ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                                    subj.send_keys(Keys.DELETE)
                                except Exception:
                                    subj.clear()

                                # fetch existing history
                                try:
                                    history_ta = WebDriverWait(driver, 3).until(
                                        EC.presence_of_element_located((
                                            By.CSS_SELECTOR, "textarea[formcontrolname='messageHistory']"
                                        ))
                                    )
                                    message_history = history_ta.get_attribute("value").strip()
                                except TimeoutException:
                                    message_history = ""

                                # type your reply (signed with CXPC name)
                                ta = WebDriverWait(driver, 3).until(
                                    EC.element_to_be_clickable((
                                        By.CSS_SELECTOR, "textarea[formcontrolname='newMessageContent']"
                                    ))
                                )
                                reply_body = (
                                    f"{message_history} Please be aware you have a pickup closing soon, "
                                    "if it is a concern please let us know. Thanks"
                                )
                                reply_body = self._format_signed_message(reply_body)
                                ta.clear()
                                ta.send_keys(reply_body)

                                # click SEND
                                send_btn = WebDriverWait(driver, 3).until(
                                    EC.element_to_be_clickable((
                                        By.XPATH, "//button[.//span[normalize-space(text())='SEND']]"
                                    ))
                                )
                                send_btn.click()
                                logging.info("âœ… Sent Pickup Reminder reply")
                                if self.log_callback:
                                    if route:
                                        self.log_callback(f"Route {route}: Pickup Reminder replied")
                                    else:
                                        self.log_callback(f"{self.dispatch_location}: Pickup Reminder replied")
                                status_str = "Pickup Reminder replied"

                                # refresh so it disappears
                                try:
                                    WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, REFRESH_BTN_CSS))
                                    ).click()
                                except TimeoutException:
                                    driver.refresh()
                                time.sleep(1)
                                try:
                                    WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.XPATH, INBOX_TAB_XPATH))
                                    ).click()
                                    time.sleep(0.5)
                                    WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))
                                    ).click()
                                except Exception as e:
                                    logging.warning(
                                        f"âš ï¸ Could not re-focus INBOX/Messages tabs: {e!r}; "
                                        "falling back to re-apply dispatch filter"
                                    )
                                    # fallback: navigate via FROâ†’Message List + redispatch
                                    try:
                                        driver.find_element(By.XPATH, "//span[text()='FRO']").click()
                                        driver.find_element(By.XPATH, "//span[text()='Message List']").click()
                                        fld = WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((
                                                By.CSS_SELECTOR,
                                                "input[formcontrolname='dispatchLocation']"
                                            ))
                                        )
                                        fld.clear()
                                        fld.send_keys(self.dispatch_location + Keys.RETURN)
                                        logging.info("ðŸ”„ Fallback: dispatch filter re-applied after reply")
                                    except Exception as ex:
                                        logging.error(f"âŒ Fallback navigation failed: {ex!r}")

                            else:
                                # â”€â”€ MOVE TO HISTORY for Pickup Reminder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                success = False
                                for attempt in range(1, 3):
                                    try:
                                        # re-find the row and history icon each time
                                        row = el.find_element(By.XPATH, "./ancestor::tr")
                                        history_icon = WebDriverWait(row, 5).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, "i.action-icon[title='Move to history']"))
                                        )
                                        # click via JS to avoid intercepts
                                        driver.execute_script("arguments[0].click();", history_icon)

                                        # wait up to 5s for that row to vanish
                                        try:
                                            WebDriverWait(driver, 5).until(staleness_of(row))
                                        except TimeoutException:
                                            logging.warning("âš ï¸ Row did not disappear after history-click; performing manual refresh")
                                            # manual fallback: re-enter dispatchLocation to reload the table
                                            try:
                                                fld = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")
                                                fld.click()
                                                fld.clear()
                                                fld.send_keys(self.dispatch_location + Keys.RETURN)
                                                # wait again for the old row to go stale
                                                WebDriverWait(driver, 5).until(staleness_of(row))
                                            except Exception as e:
                                                logging.error(f"âŒ Fallback refresh failed: {e!r}")

                                        status_str = "Pickup Reminder moved to history"
                                        if self.log_callback:
                                            if route:
                                                self.log_callback(f"Route {route}: {status_str}")
                                            else:
                                                self.log_callback(f"{self.dispatch_location}: {status_str}")
                                        success = True
                                        break

                                    except Exception as e:
                                        logging.warning(f"âš ï¸ Move to history attempt {attempt} failed: {e!r}")
                                        time.sleep(0.5)

                                if not success:
                                    logging.error("âŒ Could not move Pickup Reminder to history after 2 attempts")
                                    continue
                                # small pause before next poll
                                time.sleep(0.5)  # Increased delay to prevent crashes

                        # â”€â”€â”€ handle Early PU exactly like Pickup Reminder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        elif mtype == "Early PU" and self.early_pu_reply:
                            # Extract route number before opening reply
                            route = None
                            try:
                                cells = el.find_elements(By.TAG_NAME, "td")
                                header_ths = driver.find_elements(By.CSS_SELECTOR, "table thead tr th")
                                from_idx = next(
                                    (i for i, th in enumerate(header_ths)
                                    if "from" in th.text.strip().lower()),
                                    None
                                )
                                if from_idx is not None and from_idx < len(cells):
                                    raw_route = cells[from_idx].text.strip()
                                    m_route = re.search(r"\d{1,4}", raw_route)
                                    route = m_route.group(0) if m_route else None
                            except Exception as e:
                                logging.warning(f"âš ï¸ Could not extract route for Early PU: {e}")
                            
                            try:
                                ActionChains(driver).double_click(el).perform()
                            except:
                                el.click()
                            subj = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, 'input[formcontrolname="newMessageSubject"]'
                                ))
                            )
                            try:
                                subj.click()
                                ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                                subj.send_keys(Keys.DELETE)
                            except:
                                subj.clear()
                            ta = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, 'textarea[formcontrolname="newMessageContent"]'
                                ))
                            )
                            ta.clear()
                            reply_body = (
                                "Early pick ups are not allowed! Contact dispatch if you need the ready time "
                                "changed next time! Early pick ups are being monitored and reported! Thanks!"
                            )
                            reply_body = self._format_signed_message(reply_body)
                            ta.send_keys(reply_body)
                            send_btn = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((
                                    By.XPATH, "//button[.//span[normalize-space(text())='SEND']]"
                                ))
                            )
                            send_btn.click()
                            logging.info("Sent Early PU reply")
                            if self.log_callback:
                                if route:
                                    self.log_callback(f"Route {route}: Early PU replied")
                                else:
                                    self.log_callback(f"{self.dispatch_location}: Early PU replied")
                            status_str = "Early PU replied"
                            try:
                                WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, REFRESH_BTN_CSS))
                                ).click()
                            except TimeoutException:
                                driver.refresh()
                            time.sleep(1)
                            driver.find_element(By.XPATH, INBOX_TAB_XPATH).click()
                            time.sleep(0.5)
                            driver.find_element(By.XPATH, MSG_TAB_XPATH).click()
                            time.sleep(0.5)  # Additional delay to prevent crashes

                        # â”€â”€â”€ move everything else to history â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        else:
                            try:
                                row = el.find_element(By.XPATH, "./ancestor::tr")
                                history_icon = row.find_element(
                                    By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                driver.execute_script("arguments[0].scrollIntoView(true);", history_icon)
                                time.sleep(0.2)
                                driver.execute_script("arguments[0].click();", history_icon)
                                status_str = f"{mtype} moved to history"
                                if self.log_callback and mtype in ("Early PU", "No Pickup List"):
                                    self.log_callback(f"{self.dispatch_location}: {mtype} moved to history")
                            except Exception:
                                logging.exception(f"âŒ Failed to move {mtype} to history")
                                continue
                            time.sleep(0.8)  # Increased delay to prevent crashes

                        if self.update_status:
                            self.update_status(self.dispatch_location, status_str)

                        action_taken = True
                        
                        # â”€â”€â”€ ADD SMALL DELAY AND CLEANUP TO PREVENT CRASHES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        time.sleep(0.5)  # Small delay between message types
                        
                        # Clean up any stale elements or overlays
                        try:
                            # Close any open reply panes or overlays
                            close_btns = driver.find_elements(By.XPATH, "//span[text()='CLOSE']")
                            for close_btn in close_btns:
                                if close_btn.is_displayed():
                                    close_btn.click()
                                    time.sleep(0.2)
                        except:
                            pass
                        
                        # Small page refresh to prevent memory buildup
                        try:
                            driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                            time.sleep(1)
                        except:
                            pass
                        
                        break  # break out of for loop

                    # â”€â”€â”€ Check INBOX tab if no action taken in SYS MSGS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if not action_taken:
                        try:
                            click_tab(INBOX_TAB_XPATH, "INBOX")
                            time.sleep(0.5)
                        except Exception as e:
                            logging.warning(f"âš ï¸ Could not switch to INBOX tab: {e}")
                            # Try refreshing the page and retrying
                            try:
                                driver.refresh()
                                time.sleep(3)
                                # Wait for page to load and try again
                                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                                continue
                            except Exception as refresh_error:
                                logging.error(f"âŒ Failed to refresh page: {refresh_error}")
                                break
                        
                        # â”€â”€â”€ Auto-WalkUp detection in INBOX tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        logging.info(f"ðŸ” Auto-WalkUp enabled: {self.auto_walkup}")
                        if self.auto_walkup:
                            try:
                                # 0) refresh the list so new WalkUps show up immediately
                                try:
                                    driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                    # wait for any overlay to clear
                                    WebDriverWait(driver, 3).until(
                                        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-widget-overlay"))
                                    )
                                except Exception:
                                    pass
                            except Exception:
                                pass
                            # 1) locate non-empty "#wu" spans and snapshot their text
                            xpath = (
                                "//span["
                                "contains(translate(normalize-space(.),"
                                " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                                " 'abcdefghijklmnopqrstuvwxyz'), '#wu')"
                                " or contains(translate(normalize-space(.),"
                                " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                                " 'abcdefghijklmnopqrstuvwxyz'), '#walkup')"
                                " or contains(translate(normalize-space(.),"
                                " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                                " 'abcdefghijklmnopqrstuvwxyz'), '#walk up')"
                                "]"
                            )
                            try:
                                raw_spans = driver.find_elements(By.XPATH, xpath)
                                logging.info(f"ðŸ” Found {len(raw_spans)} potential WalkUp spans in INBOX")
                            except Exception as e:
                                logging.warning(f"âš ï¸ Failed to find WalkUp spans in INBOX: {e}")
                                raw_spans = []
                            
                            # Pair each element with its initial text, so we can fall back if it goes stale
                            walkup_data = []
                            for s in raw_spans:
                                try:
                                    if s.text.strip():
                                        # Get the text first before any operations
                                        preview_text = s.text.strip()
                                        
                                        # Get route information BEFORE opening the message (to avoid stale elements)
                                        try:
                                            row = s.find_element(By.XPATH, "./ancestor::tr")
                                            cells = row.find_elements(By.TAG_NAME, "td")
                                            
                                            # Find the "From" column index
                                            header_ths = driver.find_elements(By.CSS_SELECTOR, "table thead tr th")
                                            from_idx = next(
                                                (i for i, th in enumerate(header_ths)
                                                if "from" in th.text.strip().lower()),
                                                None
                                            )
                                            if from_idx is not None:
                                                raw_route = cells[from_idx].text.strip()
                                                m_route = re.search(r"\d{1,4}", raw_route)
                                                route = m_route.group(0) if m_route else None
                                                logging.info(f"ðŸ” Found route {route} for walkup message in INBOX")
                                            else:
                                                route = None
                                                logging.warning("âš ï¸ Could not find 'From' column header")
                                        except Exception as e:
                                            route = None
                                            logging.warning(f"âš ï¸ Could not get route from row: {e}")
                                        
                                        try:
                                            # Double-click the message row to open it and get full text
                                            row = s.find_element(By.XPATH, "./ancestor::tr")
                                            ActionChains(driver).double_click(row).perform()
                                            time.sleep(1.0)  # Wait for message to open
                                            
                                            try:
                                                # Look for the td element with the full message text
                                                full_text_td = driver.find_element(By.CSS_SELECTOR, "td[style*='white-space: pre']")
                                                if full_text_td.is_displayed():
                                                    full_msg = full_text_td.text.strip()
                                                    logging.info(f"ðŸ” Got full text from opened message in INBOX: {full_msg}")
                                                else:
                                                    full_msg = preview_text
                                                    logging.info(f"ðŸ” Full text td not visible, using preview in INBOX: {full_msg}")
                                            except:
                                                full_msg = preview_text
                                                logging.info(f"ðŸ” No full text td found, using preview in INBOX: {full_msg}")
                                            
                                            # Process the walkup immediately while the message is still open
                                            walkup_data.append((s, full_msg, route))
                                            logging.info("ðŸ” Added walkup data for processing")
                                            
                                            # Click CLOSE button to return to inbox
                                            try:
                                                close_btn = driver.find_element(By.XPATH, "//span[@class='p-button-label' and text()='CLOSE']")
                                                close_btn.click()
                                                time.sleep(0.5)  # Wait for inbox to reload
                                                logging.info("ðŸ” Clicked CLOSE button to return to inbox")
                                            except Exception as e:
                                                logging.warning(f"âš ï¸ Could not click CLOSE button: {e}")
                                        except Exception as e:
                                            logging.warning(f"âš ï¸ Could not get full message for INBOX span: {e}")
                                            walkup_data.append((s, preview_text, route))
                                except Exception as e:
                                    logging.warning(f"âš ï¸ Failed to get INBOX span text: {e}")
                                    continue
                            
                            logging.info(f"ðŸ” WalkUp data in INBOX: {[text for _, text, _ in walkup_data]}")

                            # Process the walkup data we collected
                            for span, full_msg, route in walkup_data:
                                # We already have the full message and route from the double-click extraction
                                # No need to hover or re-locate elements
                                
                                if not route:
                                    logging.warning("âš ï¸ No route found for walkup message, skipping.")
                                    continue

                                # 5) ZIP (5 digits)
                                m_zip    = re.search(r"\b(\d{5})\b", full_msg)
                                zip_code = m_zip.group(1) if m_zip else ""
                            
                                # 5) Validate ZIP code before proceeding
                                if not zip_code or not re.fullmatch(r"\d{5}", zip_code):
                                    logging.warning(f"âš ï¸ Invalid ZIP code ({zip_code!r}) for route {route} - skipping WalkUp creation")
                                    if self.log_callback:
                                        self.log_callback(f"Route {route}: WalkUp creation skipped - invalid ZIP code")
                                    # Archive the message without creating WalkUp
                                    try:
                                        row = span.find_element(By.XPATH, "./ancestor::tr")
                                        icon = row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                        icon.click()
                                        logging.info(f"âœ… Archived WalkUp message with invalid ZIP for route {route}")
                                    except Exception as e:
                                        logging.warning(f"âš ï¸ Failed to archive WalkUp message with invalid ZIP: {e!r}")
                                    continue

                                # 6) parse address from the message (using messageclear2 pattern)
                                # strip off "#wu" and trailing ZIP to get raw street portion
                                raw_addr = re.sub(r"#wu", "", full_msg, flags=re.IGNORECASE).strip()
                                # Extract ZIP only if it appears at end-of-string
                                m_zip_end = re.search(r"(\b\d{5}\b)\s*$", full_msg)
                                if m_zip_end:
                                    raw_addr = raw_addr[: m_zip_end.start()].strip()
                                
                                # 6b) Normalize address with geocoding from the start, then parse
                                address = None
                                try:
                                    # Use geocoding as primary method for consistent normalization
                                    normalized = normalize_with_geocode(raw_addr, zip_code)
                                    logging.info(f"âœ… Geocoded address: {normalized}")

                                    # Parse the normalized address
                                    tagged, _ = usaddress.tag(normalized)
                                    if tagged.get("AddressNumber") and tagged.get("StreetName"):
                                        num = tagged.get("AddressNumber", "")
                                        pre_dir = tagged.get("StreetNamePreDirectional", "")
                                        street = tagged.get("StreetName", "")
                                        suffix = tagged.get("StreetNamePostType", "")
                                        post_dir = tagged.get("StreetNamePostDirectional", "")
                                        address = " ".join(
                                            p for p in (num, pre_dir, street, suffix, post_dir) if p
                                        ).strip()
                                        logging.info(f"âœ… Standardized address: {address}")
                                    else:
                                        logging.warning("âš ï¸ Can't parse geocoded address: defaulting to 'Walk Up'")
                                        address = "Walk Up"
                                except Exception as e:
                                    logging.warning(
                                        f"âš ï¸ Geocoding failed: {e!r}; defaulting to 'Walk Up'"
                                    )
                                    address = "Walk Up"

                                # 7) Validate ZIP code after geocoding but before creating WalkUp
                                if not zip_code or not re.fullmatch(r"\d{5}", zip_code):
                                    logging.warning(f"âš ï¸ Invalid ZIP code ({zip_code!r}) for route {route} - skipping WalkUp creation")
                                    if self.log_callback:
                                        self.log_callback(f"Route {route}: WalkUp creation skipped - invalid ZIP code")
                                    # Archive the message without creating WalkUp
                                    try:
                                        row = span.find_element(By.XPATH, "./ancestor::tr")
                                        icon = row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                        icon.click()
                                        logging.info(f"âœ… Archived WalkUp message with invalid ZIP for route {route}")
                                    except Exception as e:
                                        logging.warning(f"âš ï¸ Failed to archive WalkUp message with invalid ZIP: {e!r}")
                                    continue
                                
                                # 8) create the WalkUp pickup request
                                logging.info(f"ðŸš¶ Creating WalkUp for route {route}: {address}, {zip_code}")
                                if self.update_status:
                                    self.update_status(self.dispatch_location, f"Creating WalkUp for route {route}")
                                
                                # We already have the full message content, so we can create the walkup directly
                                # Use a default stop ID since we already extracted the full message
                                stop_id = "WALKUP"
                                
                                # create the WalkUp pickup request
                                try:
                                    self._create_walkup_pickup(driver, route, address, zip_code, stop_id)
                                except Exception as e:
                                    if "VALIDATION_MISMATCH:" in str(e):
                                        # Extract the validated location from the exception
                                        validated_loc = str(e).split(":")[1]
                                        logging.warning(f"âš ï¸ ZIP validation mismatch - sending reply to route {route}")
                                        
                                        # Send reply message about ZIP mismatch
                                        try:
                                            # Click the existing Messages tab (not create new one)
                                            wait = WebDriverWait(driver, 10)
                                            messages_tab = wait.until(EC.element_to_be_clickable(
                                                (By.XPATH, "//span[contains(@class, 'p-tabview-title') and contains(text(), 'Messages - HWOA')]")))
                                            messages_tab.click()
                                            time.sleep(1)
                                            
                                            # Find the original message row
                                            message_xpath = f"//span[contains(normalize-space(.), '{preview_text}')]/ancestor::tr"
                                            message_row = wait.until(EC.presence_of_element_located((By.XPATH, message_xpath)))
                                            
                                            # Double-click to open reply
                                            ActionChains(driver).double_click(message_row).perform()
                                            time.sleep(1)
                                            
                                            # Find and fill the reply textarea
                                            reply_ta = wait.until(EC.element_to_be_clickable(
                                                (By.CSS_SELECTOR, 'textarea[formcontrolname="newMessageContent"]')))
                                            reply_ta.clear()
                                            reply_body = (
                                                "Zipcode/Postal provided is for a different station. Please ensure you are using the correct zipcode/postal within the station. If you think this is an error please send a new message. Thanks."
                                            )
                                            reply_body = self._format_signed_message(reply_body)
                                            reply_ta.send_keys(reply_body)
                                            
                                            # Send the reply
                                            send_btn = wait.until(EC.element_to_be_clickable(
                                                (By.XPATH, "//button[.//span[normalize-space(text())='SEND']]")))
                                            send_btn.click()
                                            time.sleep(1)
                                            
                                            # Archive the original message
                                            icon = message_row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                            icon.click()
                                            logging.info(f"âœ… Sent ZIP mismatch reply to route {route} and archived message")
                                            
                                            if self.log_callback:
                                                self.log_callback(f"Route {route}: ZIP validation mismatch - sent correction reply")
                                                
                                        except Exception as reply_error:
                                            logging.error(f"âŒ Failed to send ZIP mismatch reply: {reply_error}")
                                            # Try to archive the message anyway
                                            try:
                                                message_row = wait.until(EC.presence_of_element_located((By.XPATH, message_xpath)))
                                                icon = message_row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                                icon.click()
                                                logging.info(f"âœ… Archived WalkUp message for route {route} after reply failure")
                                            except:
                                                pass
                                        # Continue to next message (don't re-raise)
                                        continue
                                    else:
                                        # Re-raise other exceptions
                                        raise e
                                
                                # archive the original WalkUp message (handle stale elements)
                                try:
                                    # Re-find the message by its content since the original span might be stale
                                    message_xpath = f"//span[contains(normalize-space(.), '{preview_text}')]/ancestor::tr"
                                    message_row = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.XPATH, message_xpath))
                                    )
                                    icon = message_row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                    icon.click()
                                    logging.info(f"âœ… Archived WalkUp message for route {route}")
                                except Exception as e:
                                    logging.warning(f"âš ï¸ Failed to archive WalkUp message: {e!r}")
                                    # Try alternative approach - refresh and re-find
                                    try:
                                        driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                        time.sleep(1)
                                        message_xpath = f"//span[contains(normalize-space(.), '{preview_text}')]/ancestor::tr"
                                        message_row = WebDriverWait(driver, 5).until(
                                            EC.presence_of_element_located((By.XPATH, message_xpath))
                                        )
                                        icon = message_row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                        icon.click()
                                        logging.info(f"âœ… Archived WalkUp message for route {route} (after refresh)")
                                    except Exception as e2:
                                        logging.warning(f"âš ï¸ Failed to archive WalkUp message after refresh: {e2!r}")
                        
                        # â”€â”€â”€ Check for messages in INBOX tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        for mtype in self.message_types:
                            info = self.message_actions[mtype]

                            # â”€â”€â”€ pagination lookup â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            el = None
                            # first, ensure you're on page 1
                            try:
                                # Check if window is still valid before proceeding
                                driver.current_window_handle
                                driver.find_element(By.CSS_SELECTOR, "a.ui-paginator-first").click()
                                time.sleep(0.3)
                            except (NoSuchWindowException, NoSuchElementException):
                                # If window is closed or element not found, skip pagination
                                pass
                            except:
                                pass

                            for _ in range(3):  # try pages 1â†’2â†’3
                                try:
                                    # Check if window is still valid before proceeding
                                    driver.current_window_handle
                                    el = WebDriverWait(driver, 2).until(
                                        EC.element_to_be_clickable((By.XPATH, info["xpath"]))
                                    )
                                    break
                                except (TimeoutException, StaleElementReferenceException, NoSuchWindowException) as e:
                                    # If window is closed, break out of the loop
                                    if isinstance(e, NoSuchWindowException):
                                        logging.warning("âš ï¸ Browser window closed, stopping message processing")
                                        break
                                    # click "next page" if you can
                                    try:
                                        # Check if window is still valid before proceeding
                                        driver.current_window_handle
                                        nxt = driver.find_element(By.CSS_SELECTOR, "a.ui-paginator-next")
                                        nxt.click()
                                        time.sleep(0.5)
                                    except (NoSuchWindowException, NoSuchElementException):
                                        # If window is closed or element not found, break
                                        break
                                    except:
                                        break
                            if not el:
                                continue
                            # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

                            # â”€â”€â”€ scroll into view and log â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            driver.execute_script("arguments[0].scrollIntoView(true);", el)
                            logging.info(f"ðŸ“¬ Found {mtype} in INBOX â†’ {el.text}")

                            # â”€â”€â”€ handle Pickup Reminder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            if mtype == "Pickup Reminder":
                                if self.reply_mode:
                                    # â”€â”€ RETRY OPENING THE REPLY PANE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                    subj = None
                                    for attempt in range(1, 4):
                                        try:
                                            # bring the row fully into view
                                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                                
                                            # try double-click, fallback to click/JS click
                                            try:
                                                ActionChains(driver).double_click(el).perform()
                                            except (ElementClickInterceptedException, StaleElementReferenceException):
                                                try:
                                                    el.click()
                                                except Exception:
                                                    driver.execute_script("arguments[0].click();", el)
                                
                                            # wait up to 5s for the subject field
                                            subj = WebDriverWait(driver, 5).until(
                                                EC.visibility_of_element_located((
                                                    By.CSS_SELECTOR, "input[formcontrolname='newMessageSubject']"
                                                ))
                                            )
                                            logging.info(f"âœ… Reply pane opened on attempt {attempt}")
                                            break
                                        except (TimeoutException, StaleElementReferenceException) as e:
                                            logging.warning(f"âš ï¸ Attempt {attempt}/3: reply pane not ready ({e!r})")
                                            time.sleep(1)
                                    if not subj:
                                        logging.error("âŒ Could not open Pickup Reminder reply pane after 3 attempts â€” archiving")
                                        # give the UI a quick refresh before archiving, to avoid stale state
                                        try:
                                            driver.find_element(By.CSS_SELECTOR, REFRESH_BTN_CSS).click()
                                            time.sleep(1)
                                        except:
                                            pass
                                        try:
                                            row = el.find_element(By.XPATH, "./ancestor::tr")
                                            hist_icon = WebDriverWait(row, 5).until(
                                                EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                    "i.action-icon[title='Move to history']"
                                                ))
                                            )
                                            driver.execute_script("arguments[0].click();", hist_icon)
                                            if self.update_status:
                                                self.update_status(self.dispatch_location, "Pickup Reminder moved to history")
                                        except Exception as e:
                                            logging.error(f"âš ï¸ Failed to archive after pane-open failure: {e!r}")
                                        continue  # skip to the next message

                                    # â”€â”€ NOW YOUR ORIGINAL REPLY LOGIC â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

                                    # clear any old subject text
                                    try:
                                        subj.click()
                                        ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                                        subj.send_keys(Keys.DELETE)
                                    except Exception:
                                        subj.clear()

                                    # fetch existing history
                                    try:
                                        history_ta = WebDriverWait(driver, 3).until(
                                            EC.presence_of_element_located((
                                                By.CSS_SELECTOR, "textarea[formcontrolname='messageHistory']"
                                            ))
                                        )
                                        message_history = history_ta.get_attribute("value").strip()
                                    except TimeoutException:
                                        message_history = ""

                                    # type your reply (signed with CXPC name)
                                    ta = WebDriverWait(driver, 3).until(
                                        EC.element_to_be_clickable((
                                            By.CSS_SELECTOR, "textarea[formcontrolname='newMessageContent']"
                                        ))
                                    )
                                    ta.clear()
                                    reply_body = (
                                        f"{message_history} Please be aware you have a pickup closing soon, "
                                        "if it is a concern please let us know. Thanks"
                                    )
                                    reply_body = self._format_signed_message(reply_body)
                                    ta.send_keys(reply_body)

                                    # click SEND
                                    send_btn = WebDriverWait(driver, 3).until(
                                        EC.element_to_be_clickable((
                                            By.XPATH, "//button[.//span[normalize-space(text())='SEND']]"
                                        ))
                                    )
                                    send_btn.click()
                                    logging.info("âœ… Sent Pickup Reminder reply")
                                    if self.log_callback:
                                        self.log_callback(f"{self.dispatch_location}: Pickup Reminder replied")
                                    status_str = "Pickup Reminder replied"

                                    # refresh so it disappears
                                    try:
                                        WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, REFRESH_BTN_CSS))
                                        ).click()
                                    except TimeoutException:
                                        driver.refresh()
                                    time.sleep(1)
                                    try:
                                        WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((By.XPATH, INBOX_TAB_XPATH))
                                        ).click()
                                        time.sleep(0.5)
                                        WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))
                                        ).click()
                                    except Exception as e:
                                        logging.warning(
                                            f"âš ï¸ Could not re-focus INBOX/Messages tabs: {e!r}; "
                                            "falling back to re-apply dispatch filter"
                                        )
                                        # fallback: navigate via FROâ†’Message List + redispatch
                                        try:
                                            driver.find_element(By.XPATH, "//span[text()='FRO']").click()
                                            driver.find_element(By.XPATH, "//span[text()='Message List']").click()
                                            fld = WebDriverWait(driver, 5).until(
                                                EC.element_to_be_clickable((
                                                    By.CSS_SELECTOR,
                                                    "input[formcontrolname='dispatchLocation']"
                                                ))
                                            )
                                            fld.clear()
                                            fld.send_keys(self.dispatch_location + Keys.RETURN)
                                            logging.info("ðŸ”„ Fallback: dispatch filter re-applied after reply")
                                        except Exception as ex:
                                            logging.error(f"âŒ Fallback navigation failed: {ex!r}")

                                else:
                                    # â”€â”€ MOVE TO HISTORY for Pickup Reminder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                    success = False
                                    for attempt in range(1, 3):
                                        try:
                                            # re-find the row and history icon each time
                                            row = el.find_element(By.XPATH, "./ancestor::tr")
                                            history_icon = WebDriverWait(row, 5).until(
                                                EC.element_to_be_clickable((By.CSS_SELECTOR, "i.action-icon[title='Move to history']"))
                                            )
                                            # click via JS to avoid intercepts
                                            driver.execute_script("arguments[0].click();", history_icon)

                                            # wait up to 5s for that row to vanish
                                            try:
                                                WebDriverWait(driver, 5).until(staleness_of(row))
                                            except TimeoutException:
                                                logging.warning("âš ï¸ Row did not disappear after history-click; performing manual refresh")
                                                # manual fallback: re-enter dispatchLocation to reload the table
                                                try:
                                                    fld = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")
                                                    fld.click()
                                                    fld.clear()
                                                    fld.send_keys(self.dispatch_location + Keys.RETURN)
                                                    # wait again for the old row to go stale
                                                    WebDriverWait(driver, 5).until(staleness_of(row))
                                                except Exception as e:
                                                    logging.error(f"âŒ Fallback refresh failed: {e!r}")

                                            status_str = "Pickup Reminder moved to history"
                                            if self.log_callback:
                                                self.log_callback(f"{self.dispatch_location}: {status_str}")
                                            success = True
                                            break

                                        except Exception as e:
                                            logging.warning(f"âš ï¸ Move to history attempt {attempt} failed: {e!r}")
                                            time.sleep(0.5)

                                    if not success:
                                        logging.error("âŒ Could not move Pickup Reminder to history after 2 attempts")
                                        continue
                                    # small pause before next poll
                                    time.sleep(0.2)

                            # â”€â”€â”€ handle Early PU exactly like Pickup Reminder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            elif mtype == "Early PU" and self.early_pu_reply:
                                try:
                                    ActionChains(driver).double_click(el).perform()
                                except:
                                    el.click()
                                subj = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((
                                        By.CSS_SELECTOR, 'input[formcontrolname="newMessageSubject"]'
                                    ))
                                )
                                try:
                                    subj.click()
                                    ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                                    subj.send_keys(Keys.DELETE)
                                except:
                                    subj.clear()
                                history_ta = WebDriverWait(driver, 10).until(
                                    EC.presence_of_element_located((
                                        By.CSS_SELECTOR, 'textarea[formcontrolname="messageHistory"]'
                                    ))
                                )
                                message_history = history_ta.get_attribute('value').strip()
                                ta = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((
                                        By.CSS_SELECTOR, 'textarea[formcontrolname="newMessageContent"]'
                                    ))
                                )
                                ta.clear()
                                reply_body = (
                                    "Early pick ups are not allowed! Contact dispatch if you need the ready time "
                                    "changed next time! Early pick ups are being monitored and reported! Thanks!"
                                )
                                reply_body = self._format_signed_message(reply_body)
                                ta.send_keys(reply_body)
                                send_btn = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((
                                        By.XPATH, "//button[.//span[normalize-space(text())='SEND']]"
                                    ))
                                )
                                send_btn.click()
                                logging.info("Sent Early PU reply")
                                if self.log_callback:
                                    self.log_callback(f"{self.dispatch_location}: Early PU replied")
                                status_str = "Early PU replied"
                                try:
                                    WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, REFRESH_BTN_CSS))
                                    ).click()
                                except TimeoutException:
                                    driver.refresh()
                                time.sleep(1)
                                driver.find_element(By.XPATH, INBOX_TAB_XPATH).click()
                                time.sleep(0.5)
                                driver.find_element(By.XPATH, MSG_TAB_XPATH).click()

                            # â”€â”€â”€ move everything else to history â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            else:
                                try:
                                    row = el.find_element(By.XPATH, "./ancestor::tr")
                                    history_icon = row.find_element(
                                        By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                                    driver.execute_script("arguments[0].scrollIntoView(true);", history_icon)
                                    time.sleep(0.2)
                                    driver.execute_script("arguments[0].click();", history_icon)
                                    status_str = f"{mtype} moved to history"
                                    
                                    # Extract route information for logging
                                    route = None
                                    if self.log_callback and mtype in ("Early PU", "No Pickup List"):
                                        try:
                                            cells = row.find_elements(By.TAG_NAME, "td")
                                            header_ths = driver.find_elements(By.CSS_SELECTOR, "table thead tr th")
                                            from_idx = next(
                                                (i for i, th in enumerate(header_ths)
                                                if "from" in th.text.strip().lower()),
                                                None
                                            )
                                            if from_idx is not None and from_idx < len(cells):
                                                raw_route = cells[from_idx].text.strip()
                                                m_route = re.search(r"\d{1,4}", raw_route)
                                                route = m_route.group(0) if m_route else None
                                        except Exception as e:
                                            logging.warning(f"âš ï¸ Could not extract route for {mtype}: {e}")
                                        
                                        if route:
                                            self.log_callback(f"Route {route}: {mtype} moved to history")
                                        else:
                                            self.log_callback(f"{self.dispatch_location}: {mtype} moved to history")
                                except Exception:
                                    logging.exception(f"âŒ Failed to move {mtype} to history")
                                    continue
                                time.sleep(0.5)

                            if self.update_status:
                                self.update_status(self.dispatch_location, status_str)

                            action_taken = True
                            break  # break out of for loop

                    # â”€â”€â”€ If no action taken in either tab, go to idle state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if not action_taken:
                        logging.info("ðŸ” No messages in SYS MSGS or INBOX â€“ idle")
                        

                        
                        if self.update_status:
                            self.update_status(self.dispatch_location, IDLE_STATUS_TEXT)
                        idle_sleep = 2 if (self.reply_mode and "Pickup Reminder" in self.message_types) else POLL_FREQ
                        if self.stop_event.wait(idle_sleep):
                            break

                # normal exit inner loop â†’ break retry loop
                break

            except (ProtocolError, RemoteDisconnected):
                break
            except Exception:
                logging.exception("âŒ Auto-Read monitor error, restarting in 5s")
                if self.update_status:
                    self.update_status(self.dispatch_location, "Error â€“ restartingâ€¦")
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
                if self.stop_event.wait(5):
                    break  # stop requested
                continue  # retry

        # final cleanup & callback
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass
        logging.info(f"ðŸ›‘ Auto Read monitor for {self.dispatch_location} stopped")
        if self.finish_callback:
            self.finish_callback(self.dispatch_location)



    def _send_message(self, driver, route, text):
        """Send a message reply to a specific route"""
        wait = WebDriverWait(driver, 10)
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
        except TimeoutException:
            pass
        row_xpath = f"//td[contains(.,'{route}') and contains(.,'WORKAREA')]"
        row = wait.until(EC.element_to_be_clickable((By.XPATH, row_xpath)))
        ActionChains(driver).double_click(row).perform()
        ta = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'textarea[formcontrolname="newMessageContent"]')))
        ta.clear(); ta.send_keys(text)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[normalize-space(text())='SEND']"))).click()
        logging.info(f"âœ‰ï¸ Sent reply: {text!r}")

    def _create_walkup_pickup(self, driver, route, address, zip_code, stop_id):
        """Create a WalkUp pickup request"""
        wait = WebDriverWait(driver, 10)
        try:
            # First, make sure we're on the home page
            try:
                # Try to find and click home button if we're not already there
                home_btn = driver.find_element(By.XPATH, "//span[text()='Home']")
                if home_btn.is_displayed():
                    home_btn.click()
                    time.sleep(1)
            except:
                pass  # We might already be on home page
            
            # Navigate to Pickup Request (check if tab already exists first)
            logging.info(f"ðŸ”€ Navigating to Pickup Request form...")
            try:
                # First check if Pickup Request tab is already open
                pickup_tab_xpath = "//span[contains(@class, 'p-tabview-title') and contains(text(), 'Pickup Request')]"
                existing_tabs = driver.find_elements(By.XPATH, pickup_tab_xpath)
                
                if existing_tabs:
                    # Pickup Request tab already exists, just click on it
                    logging.info("âœ… Pickup Request tab already exists, switching to it")
                    existing_tabs[0].click()
                    time.sleep(1)
                else:
                    # No existing tab, create new one via FRO menu
                    logging.info("ðŸ“ Creating new Pickup Request tab")
                    try:
                        # Try FRO menu first
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']"))).click()
                        time.sleep(0.5)
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Pickup Request')]"))).click()
                        time.sleep(1)
                    except Exception as e:
                        logging.warning(f"âš ï¸ FRO menu navigation failed: {e}")
                        # Fallback: try direct Pickup Request button
                        try:
                            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Pickup Request')]"))).click()
                            time.sleep(1)
                        except Exception as e2:
                            logging.error(f"âŒ Direct Pickup Request navigation also failed: {e2}")
                            raise
            except Exception as e:
                logging.error(f"âŒ Failed to navigate to Pickup Request: {e}")
                raise
            
            logging.info(f"ðŸ“ Filling WalkUp form for route {route}...")
            
            # Wait for form to load
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                logging.info("âœ… Pickup Request form loaded successfully")
            except Exception as e:
                logging.error(f"âŒ Pickup Request form failed to load: {e}")
                raise
            
            # Validate Services BEFORE opening Quick Pickup popup
            logging.info("ðŸ” Validating services...")
            validation_passed = False
            try:
                # Click "Validate Services" button
                validate_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@label='Validate Services' and contains(@class,'pickupButtons')]")))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", validate_btn)
                try:
                    validate_btn.click()
                except ElementClickInterceptedException:
                    driver.execute_script("arguments[0].click();", validate_btn)
                
                # Enter ZIP into postalSvcValid
                svc_in = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[formcontrolname='postalSvcValid']")))
                svc_in.clear()
                svc_in.send_keys(zip_code)
                
                # Click the Validate button
                validate_submit_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(@class,'pickupButtons') and normalize-space(text())='Validate']")))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", validate_submit_btn)
                try:
                    validate_submit_btn.click()
                except ElementClickInterceptedException:
                    driver.execute_script("arguments[0].click();", validate_submit_btn)
                
                # Wait for Service Validation popup
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((
                        By.XPATH, "//span[contains(normalize-space(.),'Service Validation')]"
                    ))
                )
                
                # Wait for validated location
                def non_empty_value(drv):
                    inp = drv.find_element(By.CSS_SELECTOR, "input[formcontrolname='locSvcValid']")
                    val = inp.get_attribute("value").strip()
                    return val or False
                
                validated_loc = WebDriverWait(driver, 10).until(non_empty_value).upper()
                logging.info(f"ðŸ” Service-validated location: {validated_loc}")
                
                # Check for validation mismatch
                if validated_loc != self.dispatch_location.upper():
                    logging.warning(f"âš ï¸ Validation mismatch: expected='{self.dispatch_location.upper()}', got='{validated_loc}'")
                    # Close the Service Validation popup
                    close_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(@class,'pickupButtons') and normalize-space(text())='Close']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", close_btn)
                    try:
                        close_btn.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", close_btn)
                    logging.info("âœ… Service Validation popup closed due to mismatch")
                    # Return to message list to send reply
                    raise Exception(f"VALIDATION_MISMATCH:{validated_loc}")
                else:
                    # Close the Service Validation popup
                    close_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(@class,'pickupButtons') and normalize-space(text())='Close']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", close_btn)
                    try:
                        close_btn.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", close_btn)
                    logging.info("âœ… Service validation passed")
                    validation_passed = True
                
            except Exception as e:
                if "VALIDATION_MISMATCH:" in str(e):
                    # Re-raise validation mismatch exceptions
                    logging.info("ðŸ›‘ Validation mismatch detected - stopping WalkUp creation")
                    raise e
                else:
                    logging.warning(f"âš ï¸ Service validation failed: {e}")
                    # Continue with submission even if validation fails
                    validation_passed = True
            
            # Only proceed with WalkUp creation if validation passed or failed (but not mismatched)
            if not validation_passed:
                logging.info("ðŸ›‘ Stopping WalkUp creation due to validation mismatch")
                return
            
            # Check if Quick Pickup popup is already open, if not open it
            logging.info("ðŸ”˜ Checking Quick Pickup popup...")
            try:
                # Check if Quick Pickup popup is already open
                existing_popup = driver.find_elements(By.CSS_SELECTOR, "input[formcontrolname='walkupAddress']")
                if existing_popup:
                    logging.info("âœ… Quick Pickup popup already open")
                else:
                    # Open Quick Pickup popup
                    logging.info("ðŸ”˜ Opening Quick Pickup popup...")
                    quick_pickup_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(text(), 'Quick Pickup')]")))
                    quick_pickup_btn.click()
                    time.sleep(1)
                    logging.info("âœ… Quick Pickup popup opened")
            except Exception as e:
                logging.error(f"âŒ Failed to handle Quick Pickup popup: {e}")
                raise
            
            # Wait for Quick Pickup popup to load
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='walkupAddress']")))
                logging.info("âœ… Quick Pickup popup form loaded")
            except Exception as e:
                logging.error(f"âŒ Quick Pickup popup failed to load: {e}")
                raise
            
            # Fill in the Quick Pickup popup form
            # Address
            addr_field = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[formcontrolname='walkupAddress']")))
            addr_field.clear()
            addr_field.send_keys(address)
            time.sleep(0.3)
            
            # ZIP Code
            zip_field = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[formcontrolname='walkupPostal']")))
            zip_field.clear()
            zip_field.send_keys(zip_code)
            time.sleep(0.3)
            
            # Route
            route_field = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[formcontrolname='walkupRoute']")))
            route_field.clear()
            route_field.send_keys(str(route))
            time.sleep(0.3)
            
            # Submit
            logging.info(f"ðŸš€ Submitting WalkUp request...")
            submit_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'pickupButtons') and normalize-space(text())='Submit']")))
            submit_btn.click()
            time.sleep(1)
            
            # Handle route conflict dialog if it appears
            try:
                logging.info("ðŸ” Checking for route conflict dialog...")
                yes_btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(@class,'pickupButtons') and normalize-space(text())='Yes']")))
                yes_btn.click()
                logging.info("âœ… Clicked 'Yes' for route conflict")
                time.sleep(1)
            except:
                logging.info("â„¹ï¸ No route conflict dialog found")
            
            logging.info(f"âœ… WalkUp pickup request created for route {route}")
            
            # Close any "Details" tabs that were created during the pickup process and extract pickup ID
            pickup_id = None
            try:
                logging.info("ðŸ” Looking for Details tabs to close...")
                details_tabs = driver.find_elements(By.XPATH, "//span[contains(@class, 'p-tabview-title') and contains(text(), 'Details')]")
                for details_tab in details_tabs:
                    try:
                        tab_title = details_tab.text.strip()
                        logging.info(f"ðŸ” Found Details tab: {tab_title}")
                        
                        # Extract pickup ID from tab title (format: "Details HWOA/3182/08-07")
                        if "Details" in tab_title and "/" in tab_title:
                            parts = tab_title.split("/")
                            if len(parts) >= 2:
                                pickup_id = parts[1].strip()
                                logging.info(f"ðŸ” Extracted pickup ID: {pickup_id}")
                        
                        # Find the close button (X) for this Details tab
                        tab_element = details_tab.find_element(By.XPATH, "./ancestor::a[contains(@class, 'p-tabview-nav-link')]")
                        close_button = tab_element.find_element(By.CSS_SELECTOR, "svg.p-icon.p-tabview-close")
                        close_button.click()
                        time.sleep(0.5)
                        logging.info(f"âœ… Closed Details tab: {tab_title}")
                    except Exception as close_error:
                        logging.warning(f"âš ï¸ Failed to close Details tab: {close_error}")
            except Exception as e:
                logging.warning(f"âš ï¸ Failed to find Details tabs: {e}")
            
            # Log the WalkUp creation with pickup ID and location if available
            if self.log_callback:
                if pickup_id:
                    self.log_callback(f"Route {route}: WalkUp pickup request created (ID: {pickup_id}) at {address}, {zip_code}")
                else:
                    self.log_callback(f"Route {route}: WalkUp pickup request created at {address}, {zip_code}")
            
            # Don't close the Quick Pickup popup - keep it open for future use
            logging.info("âœ… WalkUp pickup request completed - keeping tabs open")
            
            # Navigate back to Message List to archive the original message
            try:
                logging.info("ðŸ”€ Returning to Message List...")
                # Try to click on Messages tab directly first
                try:
                    msg_tab = wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH)))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", msg_tab)
                    msg_tab.click()
                    time.sleep(1)
                    logging.info("âœ… Returned to Message List via tab click")
                except Exception as tab_error:
                    logging.warning(f"âš ï¸ Direct tab click failed: {tab_error}")
                    # Fallback: use FRO menu navigation
                    try:
                        fro_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']")))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", fro_btn)
                        fro_btn.click()
                        time.sleep(0.5)
                        msg_list_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Message List']")))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", msg_list_btn)
                        msg_list_btn.click()
                        time.sleep(1)
                        logging.info("âœ… Returned to Message List via FRO menu")
                    except Exception as fro_error:
                        logging.warning(f"âš ï¸ FRO menu navigation failed: {fro_error}")
                        # Final fallback: try to refresh and navigate
                        driver.refresh()
                        time.sleep(2)
                        wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
                        time.sleep(1)
                        logging.info("âœ… Returned to Message List via refresh")
            except Exception as e:
                logging.warning(f"âš ï¸ Failed to return to Message List: {e}")
                
        except Exception as e:
            # Check if this is a validation mismatch exception - if so, re-raise it
            if "VALIDATION_MISMATCH:" in str(e):
                raise e  # Re-raise validation mismatch exceptions to be handled by the caller
            
            logging.error(f"âŒ Failed to create WalkUp pickup request: {type(e).__name__}: {str(e)}")
            if self.log_callback:
                self.log_callback(f"{self.dispatch_location}: WalkUp creation failed for route {route}")
            # Try to recover by returning to Message List
            try:
                # Navigate back to Message List to continue monitoring
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']"))).click()
                time.sleep(0.5)
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Message List']"))).click()
                time.sleep(1)
                logging.info("âœ… Recovered to Message List after WalkUp creation failure")
            except Exception as recovery_error:
                logging.error(f"âŒ Failed to recover to Message List: {recovery_error}")
                # If recovery fails, try refreshing the page
                try:
                    driver.refresh()
                    time.sleep(3)
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                    logging.info("âœ… Page refreshed after WalkUp creation failure")
                except Exception as refresh_error:
                    logging.error(f"âŒ Failed to refresh page after WalkUp creation failure: {refresh_error}")

# â”€â”€â”€ Auto-Clear Monitor with automatic restart â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class CarrierClearSeleniumMonitor:
    def __init__(self, username, password, dispatch_location,
                 clear_reply, stop_event, skip_routes=None,
                 update_status=None, finish_callback=None, log_callback=None,
                 mfa_factor_callback=None, mfa_code_callback=None,
                 agent_display_name=None):
        self.username             = username
        self.password             = password
        self.dispatch_location    = dispatch_location.upper()
        self.clear_reply          = clear_reply
        self.agent_display_name   = agent_display_name
        self.stop_event           = stop_event
        self.skip_routes          = skip_routes if skip_routes is not None else set()
        self.update_status        = update_status
        self.finish_callback      = finish_callback
        self.log_callback      = log_callback
        self.mfa_factor_callback  = mfa_factor_callback
        self.mfa_code_callback    = mfa_code_callback
        self.driver               = None
        self.cleared_routes       = set()
        self.open_routes_reported = set()
        self.history              = []

    def _format_clear_reply(self) -> str:
        """
        Build the final clear message that goes back to the courier.
        Appends "-CXPC First L." when we know the dispatcher name.
        """
        base = (self.clear_reply or "Clear").strip()
        name = (self.agent_display_name or "").strip()
        if not name:
            return base
        # Avoid double-appending if base already ends with this signature
        signature = f"-CXPC {name}"
        if base.endswith(signature) or signature in base:
            return base
        return f"{base} {signature}"

    def _sign_text(self, text: str) -> str:
        """
        Append the CXPC signature to an arbitrary outbound message body
        (used for open-stop warnings, Skip Clear replies, etc.).
        """
        base = (text or "").rstrip()
        name = (self.agent_display_name or "").strip()
        if not name:
            return base
        signature = f"-CXPC {name}"
        if base.endswith(signature) or signature in base:
            return base
        return f"{base} {signature}"

    def _send_clear_via_create_message(self, driver, route: str, text: str):
        """
        Compose and send a brand-new clear message to a route using the
        CREATE MESSAGE tab. This does NOT rely on an existing WORKAREA row.
        """
        wait = WebDriverWait(driver, 10)

        # 1) Switch to CREATE MESSAGE tab
        wait.until(EC.element_to_be_clickable((By.XPATH, CREATE_MSG_TAB_XPATH))).click()

        # 2) Filter for the route in the source picklist
        filt = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "input.p-picklist-filter-input[data-pc-section='sourceFilterInput']",
                )
            )
        )
        filt.clear()
        filt.send_keys(route)
        time.sleep(0.3)

        # 3) Click the matching route entry
        route_div = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    f"//div[contains(@class,'p-helper-clearfix') and normalize-space(text())='{route}']",
                )
            )
        )
        route_div.click()

        # 4) Move it to the target side
        move_btn = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[data-pc-section='moveToTargetButton']")
            )
        )
        move_btn.click()

        # 5) Type the message body
        ta = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "textarea[formcontrolname='newMessageContent']")
            )
        )
        ta.clear()
        ta.send_keys(text)

        # 6) Click SEND
        send_btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[.//span[normalize-space(text())='SEND']]")
            )
        )
        send_btn.click()
        logging.info(f"âœ‰ï¸ Sent clear message via CREATE MESSAGE tab for route {route}: {text!r}")

        # 7) Return to INBOX â†’ Messages so the UI is back in its normal state
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, INBOX_TAB_XPATH))).click()
            time.sleep(0.3)
            wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
        except Exception:
            # Non-fatal if we can't navigate back; Auto-Clear will recover on next loop
            pass

    def _get_stops(self, driver):
        """
        Return a list of 4â€‘digit stop IDs that are still open.
        Ignores the synthetic "LAST STOP" summary row that was added
        to Route Summary.
        """
        for _ in range(2):
            try:
                divs = driver.find_elements(By.CSS_SELECTOR, "div.ng-star-inserted")
                stops: list[str] = []
                for d in divs:
                    txt = d.text.strip()
                    if not (txt.isdigit() and len(txt) == 4):
                        continue
                    # Skip any stop that lives in the "LAST STOP" summary row
                    try:
                        d.find_element(
                            By.XPATH,
                            "./ancestor::tr[.//span[@title='Last Stop']]",
                        )
                        # If we got here, this stop-id belongs to the LAST STOP row â†’ ignore
                        continue
                    except NoSuchElementException:
                        pass
                    stops.append(txt)
                return stops
            except StaleElementReferenceException:
                time.sleep(1)
        return []

    def _get_stop_address(self, driver, stop_id: str) -> str:
        """
        Best-effort helper to fetch the street address for a given stop ID
        from the current Route Summary grid.
        """
        for _ in range(2):
            try:
                # Find the table row that contains this stop id
                row = driver.find_element(
                    By.XPATH,
                    f"//tr[.//div[normalize-space(text())='{stop_id}']]",
                )
                # Within that row, look for a span with normal font-weight (address cell)
                spans = row.find_elements(
                    By.XPATH,
                    ".//span[contains(@style,'font-weight: normal')]",
                )
                for s in spans:
                    txt = s.text.strip()
                    if txt:
                        return txt
                return ""
            except StaleElementReferenceException:
                time.sleep(1)
            except Exception:
                return ""
        return ""

    def _send_message(self, driver, route, text):
        """Send a message reply to a specific route"""
        wait = WebDriverWait(driver, 10)
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
        except TimeoutException:
            pass
        row_xpath = f"//td[contains(.,'{route}') and contains(.,'WORKAREA')]"
        row = wait.until(EC.element_to_be_clickable((By.XPATH, row_xpath)))
        ActionChains(driver).double_click(row).perform()
        ta = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'textarea[formcontrolname="newMessageContent"]')))
        ta.clear(); ta.send_keys(text)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[normalize-space(text())='SEND']"))).click()
        logging.info(f"âœ‰ï¸ Sent reply: {text!r}")



    def run(self):
        logging.info(f"ðŸ” Starting Auto-Clear monitor for {self.dispatch_location}")
        # keep retrying until stopped
        while not self.stop_event.is_set():
            # â”€â”€ install chromedriver & silence its logs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            driver_path = ChromeDriverManager().install()
            service     = Service(driver_path, log_path=os.devnull)

            # â”€â”€ build headless-chrome options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            opts = Options()
            # Enable Chrome performance logging for Network events (to capture CSV response)
            try:
                opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
            except Exception:
                pass
            opts.add_argument("--disable-extensions")
            opts.add_argument("--headless")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--disable-gpu")

            # â”€â”€ QUIET CHROME FLAGS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            opts.add_argument("--disable-background-networking")
            opts.add_argument("--disable-gcm")
            opts.add_argument("--log-level=3")
            opts.add_experimental_option("excludeSwitches", ["enable-logging"])

            try:
                self.driver = webdriver.Chrome(service=service, options=opts)
                driver = self.driver
                wait   = WebDriverWait(driver, 20)

                # â”€â”€â”€ track consecutive refresh failures â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                refresh_fail_count = 0

                # LOGIN (only if login form appears)
                login_helper = CarrierSeleniumMonitor(
                    username=self.username,
                    password=self.password,
                    dispatch_location=self.dispatch_location,
                    stop_event=self.stop_event,
                    message_types=[],  # no message-types needed for just login
                    mfa_factor_callback=self.mfa_factor_callback,
                    mfa_code_callback  =self.mfa_code_callback
                )
                login_helper.sso_navigate(
                    driver, wait,
                    "https://internal.example.com",
                    self.username, self.password
                )

                # FAST home-page detection with up to 5 retries
                for attempt in range(5):
                    # 1) wait up to ~2â€‰s for the URL to contain '/home'
                    for _ in range(20):
                        if "/home" in driver.current_url:
                            break
                        time.sleep(0.1)
                    # 2) then try up to 1â€‰s for the FRO menu
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH,
                                "//span[normalize-space(text())='FRO']"
                            ))
                        )
                        logging.info(f"âœ… Home UI detected on attempt {attempt+1}")
                        break
                    except TimeoutException:
                        logging.warning(f"â— Home detection attempt {attempt+1} failed")
                        if attempt < 4:
                            logging.info("ðŸ”„ Refreshing page and retryingâ€¦")
                            driver.refresh()
                            time.sleep(7)
                        else:
                            logging.error("âŒ DSM header never appeared; aborting Auto-Clear")
                            return

                # GO TO MESSAGES
                logging.info("âž¡ï¸  Switching to Messages tab")

                # â€¦then go to Messages via the menu (not the old tab XPath)â€¦
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Message List']"))).click()

                # â€¦and re-apply your dispatch-location filter
                fld = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                fld.clear()
                fld.send_keys(self.dispatch_location + Keys.RETURN)
                time.sleep(1)
                fld = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                fld.clear()
                fld.send_keys(self.dispatch_location + Keys.RETURN)
                logging.info(f"âœ… Dispatch location set to {self.dispatch_location} (Auto-Clear)")

                # PRE-OPEN ROUTE SUMMARY & BACK
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Route Summary']"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()

                # initial idle
                if self.update_status:
                    self.update_status(self.dispatch_location, CLEAR_IDLE_STATUS_TEXT)
                logging.info("â„¹ï¸  Auto-Clear idle, waiting for new requestsâ€¦")

                # MONITOR LOOP
                while not self.stop_event.is_set():
                    # idle indicator
                    if self.update_status:
                        self.update_status(self.dispatch_location, CLEAR_IDLE_STATUS_TEXT)

                    # Ensure on Messages tab
                    try:
                        # Check if window is still valid before proceeding
                        driver.current_window_handle
                        wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
                        logging.info("âœ… Switched to Messages tab before refresh")
                    except (NoSuchWindowException, NoSuchElementException):
                        logging.warning("âš ï¸ Browser window closed or element not found, stopping Auto-Clear")
                        break
                    except:
                        pass

                    # Click Refresh button
                    try:
                        # Check if window is still valid before proceeding
                        driver.current_window_handle
                        refresh_btn = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, REFRESH_BTN_CSS))
                        )
                        refresh_btn.click()
                        logging.info("ðŸ”„ Refreshing message list (Auto-Clear)")
                        # reset on success
                        refresh_fail_count = 0
                    except (NoSuchWindowException, NoSuchElementException) as e:
                        logging.warning(f"âš ï¸ Browser window closed or element not found: {e}")
                        break
                    except Exception as e:
                        refresh_fail_count += 1
                        logging.warning(f"âš ï¸ Refresh attempt {refresh_fail_count}/10 failed ({e!r})")
                        if refresh_fail_count >= 10:
                            logging.error("âŒ Too many refresh failuresâ€”restarting Auto-Clear monitor")
                            raise
                        logging.warning("â€¦falling back to dispatch-box Enter")
                        try:
                            # Check if window is still valid before proceeding
                            driver.current_window_handle
                            # click into the dispatchLocation field and press Enter
                            loc_field = wait.until(
                                EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")
                                )
                            )
                            loc_field.click()
                            loc_field.send_keys(Keys.RETURN)
                            logging.info("ðŸ”„ Refreshed message list via dispatch-box Enter")
                        except (NoSuchWindowException, NoSuchElementException) as e2:
                            logging.warning(f"âš ï¸ Browser window closed or element not found during fallback: {e2}")
                            break
                        except Exception as e2:
                            logging.error(f"âŒ Fallback dispatch-box refresh failed: {e2!r}")
                        # throttle a bit before next loop
                        if self.stop_event.wait(POLL_FREQ):
                            break
                        continue

                    time.sleep(1)

                    # â”€â”€â”€ WAIT FOR NEW CLEAR-REQUEST â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    # (after refresh & a brief sleep)
                    time.sleep(1)
                    all_spans = driver.find_elements(By.XPATH, NEW_MSG_XPATH)
                    raw_cells = [span for span in all_spans if span.text.strip()]
                    logging.info(f"ðŸ” raw clear-candidates ({len(raw_cells)}): "
                                 f"{[span.text for span in raw_cells]}")

                    # Now only accept rows that explicitly say "REQUEST TO CLEAR"
                    valid = []
                    for span in raw_cells:
                        row = span.find_element(By.XPATH, "./ancestor::tr")
                        text = row.text.strip()
                        # only require "REQUEST TO" (handles the "â€¦" truncation)
                        m = re.search(r"WORKAREA\D*?(\d{1,4})\D*REQUEST TO\b", text, re.IGNORECASE)
                        if m:
                            valid.append((span, text))

                    if not valid:
                        # nothing real â†’ throttle
                        if self.stop_event.wait(POLL_FREQ):
                            break
                        continue

                    # We have at least one genuine clear-request
                    cell, text = valid[0]
                    logging.info(f"ðŸ“¬ Detected: {text}")
                    if self.update_status:
                        self.update_status(self.dispatch_location, f"Detected: {text}")

                    # â”€â”€â”€ PARSE ROUTE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    m = re.search(r"WORKAREA\D*?(\d{1,4})\D*REQUEST TO", text, re.IGNORECASE)
                    route = str(int(m.group(1)))  # safe because we filtered above
                    logging.info(f"âž¡ï¸  Route #{route} parsed")
                    if self.update_status:
                        self.update_status(self.dispatch_location, f"Route #{route} parsed")
                    # â”€â”€ Skip Clear check â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if int(route) in self.skip_routes:
                        logging.info(f"âž– Skipping clear for route {route}")
                        # send Skip Clear reply instead of clearing (signed with CXPC name)
                        skip_body = (
                            "Not Clear yet. Please contact your manager/dispatch, there are things you have to do "
                            "before clearing. Thanks"
                        )
                        skip_body = self._sign_text(skip_body)
                        self._send_message(driver, route, skip_body)
                        if self.log_callback:
                            self.log_callback(f"{self.dispatch_location}: Skip Clear reply sent for route {route}")
                        # archive the message
                        row = cell.find_element(By.XPATH, "./ancestor::tr")
                        icon = row.find_element(By.CSS_SELECTOR, "i.action-icon[title='Move to history']")
                        # archive the row
                        driver.execute_script("arguments[0].click();", icon)
                
                        # 3) mirror your duplicate-clear flow: go back to INBOX â†’ Messages
                        wait.until(EC.element_to_be_clickable((By.XPATH, INBOX_TAB_XPATH))).click()
                        wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
                
                        # 4) re-apply the dispatch-location filter so the table reloads
                        fld = wait.until(EC.element_to_be_clickable((
                            By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']"
                        )))
                        fld.clear()
                        fld.send_keys(self.dispatch_location + Keys.RETURN)
                
                        # 5) give the table a moment to refresh, then continue
                        time.sleep(1)
                        continue

                    # OPEN ROUTE SUMMARY
                    tabs = driver.find_elements(By.CSS_SELECTOR, "span.p-tabview-title")
                    if len(tabs) < 2:
                        logging.error("âŒ Route Summary tab missing")
                        if self.update_status:
                            self.update_status(self.dispatch_location, "Route Summary missing")
                        continue
                    tabs[1].click()
                    time.sleep(1)

                    # FILL & SUBMIT
                    logging.info(f"âœï¸  Filling summary {self.dispatch_location}/{route}")
                    wait.until(EC.element_to_be_clickable((By.ID, "location"))).clear()
                    wait.until(EC.element_to_be_clickable((By.ID, "location"))).send_keys(self.dispatch_location)
                    wait.until(EC.element_to_be_clickable((By.ID, "route"))).clear()
                    wait.until(EC.element_to_be_clickable((By.ID, "route"))).send_keys(route + Keys.RETURN)
                    logging.info("âœ… Summary submitted")
                    if self.update_status:
                        self.update_status(self.dispatch_location, "Summary submitted")
                    # give the page a moment to render
                    time.sleep(1)

                    # â”€â”€â”€ VERIFY ROUTE SUMMARY PAGE FULLY LOADED â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    # First, ensure the Route Summary page has fully loaded with essential elements
                    try:
                        # Try multiple table selectors
                        table_selectors = [
                            "table.table",
                            "table",
                            ".table",
                            "[class*='table']",
                            "tbody",
                            "tr"
                        ]
                        
                        table_found = False
                        for selector in table_selectors:
                            try:
                                WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                                )
                                logging.info(f"âœ… Found table using selector: {selector}")
                                table_found = True
                                break
                            except TimeoutException:
                                continue
                        
                        if not table_found:
                            # If no table found, try to find any content on the page
                            try:
                                WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                                )
                                logging.info("âœ… Page loaded, proceeding without specific table verification")
                                table_found = True
                            except TimeoutException:
                                pass
                        
                        if not table_found:
                            raise TimeoutException("No table or content found")
                        
                        # Give the page a moment to fully render
                        time.sleep(3)
                        
                        logging.info("âœ… Route Summary page loaded")
                        
                    except TimeoutException:
                        logging.error("âŒ Route Summary page failed to load; retrying in 15s")
                        # wait up to 15s (or exit if stop requested), then retry loop
                        if self.stop_event.wait(15):
                            break
                        continue

                    # â”€â”€â”€ VERIFY SPECIFIC ROUTE EXISTS IN TABLE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    # Now look for the specific route in the loaded table with multiple selectors
                    padded = route.zfill(4)
                    route_found = False
                    
                    # Try multiple XPath selectors for route detection
                    route_selectors = [
                        f"//td[span[normalize-space(text())='{route}' or normalize-space(text())='{padded}']]",
                        f"//td[contains(text(), '{route}') or contains(text(), '{padded}')]",
                        f"//span[normalize-space(text())='{route}' or normalize-space(text())='{padded}']",
                        f"//*[contains(text(), '{route}') or contains(text(), '{padded}')]"
                    ]
                    
                    for selector in route_selectors:
                        try:
                            WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, selector))
                            )
                            logging.info(f"âœ… Found route {route} in summary table using selector: {selector}")
                            route_found = True
                            break
                        except TimeoutException:
                            continue
                    
                    if not route_found:
                        logging.warning(f"ðŸš§ Route {route} not found in summary table with any selector; skipping this route")
                        continue

                    # CHECK STOPS
                    stops = self._get_stops(driver)
                    if not stops:
                        if route not in self.cleared_routes:
                            logging.info("ðŸš® No open stops â€“ clearing route")
                            logging.info("ðŸš® No open stops â€“ entering clear/block flow")

                            # 1) Find the Clear/Unclear button
                            try:
                                clear_btn = wait.until(EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, "button.clearvisible"
                                )))
                            except TimeoutException:
                                logging.error("âŒ Clear/Unclear button not found; skipping clear flow")
                                continue

                            btn_text = clear_btn.text.strip()
                            if btn_text == "Clear":
                                # â”€â”€â”€ Click Clear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                for _ in range(3):
                                    try:
                                        clear_btn.click()
                                        break
                                    except StaleElementReferenceException:
                                        time.sleep(0.5)

                                # â”€â”€â”€ Fill clearTime if empty â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                try:
                                    clear_time_input = wait.until(EC.visibility_of_element_located((
                                        By.CSS_SELECTOR, "input[formcontrolname='clearTime']"
                                    )))
                                    if not clear_time_input.get_attribute("value").strip():
                                        clear_time_input.clear()
                                        clear_time_input.send_keys(datetime.now().strftime("%H%M"))
                                except TimeoutException:
                                    logging.warning("âš ï¸ Clear-time input not found; skipping time fill")

                                # â”€â”€â”€ Confirm Clear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                ok_clear = wait.until(EC.element_to_be_clickable((
                                    By.XPATH,
                                    "//button[contains(@class,'pickupButtons') and normalize-space(text())='OK']"
                                )))
                                ok_clear.click()

                            elif btn_text == "Unclear":
                                logging.info(f"â„¹ï¸ Button text is 'Unclear' for route {route}; skipping Clear click")
                            else:
                                logging.warning(f"âš ï¸ Unexpected button text {btn_text!r}; attempting Clear anyway")
                                try:
                                    clear_btn.click()
                                except Exception as e:
                                    logging.warning(f"âš ï¸ Clear click failed: {e!r}")

                            # 2) Click Block â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            try:
                                block_btn = wait.until(EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, "button.blockvisible"
                                )))
                                block_btn.click()
                            except TimeoutException:
                                logging.warning("âš ï¸ Block button not found; skipping block step")

                            # 3) Fill blockTime if empty â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            try:
                                block_time_input = wait.until(EC.visibility_of_element_located((
                                    By.CSS_SELECTOR, "input[formcontrolname='blockTime']"
                                )))
                                if not block_time_input.get_attribute("value").strip():
                                    block_time_input.clear()
                                    block_time_input.send_keys(datetime.now().strftime("%H%M"))
                            except TimeoutException:
                                logging.warning("âš ï¸ Block-time input not found; skipping time fill")

                            # 4) Confirm Block â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            for _ in range(3):
                                try:
                                    ok_block = wait.until(EC.element_to_be_clickable((
                                        By.XPATH,
                                        "//button[contains(@class,'pickupButtons') and normalize-space(text())='OK']"
                                    )))
                                    ok_block.click()
                                    break
                                except StaleElementReferenceException:
                                    time.sleep(0.5)

                            # â”€â”€â”€ 5) Mark as cleared, log, reply, archive â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            self.cleared_routes.add(route)
                            ts = datetime.now()
                            self.history.append((route, ts))
                            logging.info(f"âœ… Route {route} cleared at {ts:%H:%M:%S}")

                            if self.log_callback:
                                self.log_callback(f"{self.dispatch_location}: Route {route} cleared at {ts:%H:%M:%S}")
                            if self.update_status:
                                self.update_status(self.dispatch_location, f"Cleared {route} at {ts:%H:%M:%S}")

                            # send your "clear" reply back to Messages (with CXPC signature when available)
                            self._send_message(driver, route, self._format_clear_reply())

                            # archive the row
                            try:
                                row  = cell.find_element(By.XPATH, "./ancestor::tr")
                                icon = row.find_element(
                                    By.CSS_SELECTOR,
                                    "i.action-icon[title='Move to history']"
                                )
                                driver.execute_script("arguments[0].click();", icon)
                            except Exception:
                                pass

                        else:
                            # â”€â”€â”€ Duplicate clear request â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            logging.warning(f"Route {route} already cleared; archiving duplicate")
                            if self.update_status:
                                self.update_status(
                                    self.dispatch_location,
                                    f"Moved duplicate clear {route} to history"
                                )
                
                            try:
                                # 1) Archive to history
                                row = cell.find_element(By.XPATH, "./ancestor::tr")
                                icon = row.find_element(
                                    By.CSS_SELECTOR,
                                    "i.action-icon[title='Move to history']"
                                )
                                driver.execute_script("arguments[0].click();", icon)
                
                                # 2) Make sure we're back on the Messages tab
                                wait.until(EC.element_to_be_clickable((By.XPATH, INBOX_TAB_XPATH))).click()
                                wait.until(EC.element_to_be_clickable((By.XPATH, MSG_TAB_XPATH))).click()
                
                                # 3) Re-set the dispatch location filter
                                fld = wait.until(EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")))
                                fld.clear()
                                fld.send_keys(self.dispatch_location + Keys.RETURN)
                            except TimeoutException:
                                # Tab-navigation took too longâ€”ignore and keep going
                                logging.info(f"â„¹ï¸  Duplicate archive nav timed out for route {route}; continuing")
                            except Exception as dup_err:
                                logging.error(f"âŒ Unexpected error archiving duplicate for route {route}: {dup_err!r}")
                
                            # small pause before re-looping so we don't spin instantly
                            time.sleep(POLL_FREQ)
                            continue
                    else:
                        logging.warning(f"âš ï¸ {len(stops)} open stops; waiting 30s")
                        if self.update_status:
                            self.update_status(self.dispatch_location,
                                               f"{len(stops)} open stops; waiting 30s")
                        time.sleep(30)
                        try:
                            wait.until(EC.element_to_be_clickable((By.XPATH, SUBMIT_BTN_XPATH))).click()
                            logging.info("ðŸ•’ Clicked Submit")
                            if self.update_status:
                                self.update_status(self.dispatch_location, "Clicked Submit")
                        except TimeoutException:
                            logging.warning("âš ï¸ Submit button not found; skipping")
                            if self.update_status:
                                self.update_status(self.dispatch_location, "Submit missing")
                        time.sleep(1)

                        stops2 = self._get_stops(driver)
                        if stops2:
                            # â”€â”€â”€ Add route to audit only if stops still exist after Submit â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            if route not in self.open_routes_reported:
                                # 1) record it FIRST (before status update so audit check can find it)
                                self.open_routes_reported.add(route)
                                logging.info(f"âœ‰ï¸ Route {route} added to audit queue (stops remain: {stops2})")

                                # 2) then update the status (which triggers _update_clear_status audit check)
                                logging.info(f"âœ‰ï¸ Reporting open stops: {stops2}")
                                if self.update_status:
                                    self.update_status(self.dispatch_location,
                                                    f"Reporting open stops: {stops2}")

                                # 3) send your reply (one line per open stop, include address when available)
                                lines = []
                                for sid in stops2:
                                    addr = self._get_stop_address(driver, sid) or ""
                                    if addr:
                                        lines.append(
                                            f"Stop ID {sid} at {addr} is still not completed. Please reply to dispatch with the time of completion. Thanks"
                                        )
                                    else:
                                        lines.append(
                                            f"Stop ID {sid} is still not completed. Please reply to dispatch with the time of completion. Thanks"
                                        )
                                body = "\n".join(lines)
                                body = self._sign_text(body)
                                self._send_message(driver, route, body)
                            else:
                                logging.info(f"âš ï¸ Still {len(stops2)} open stops after submit (already reported)")
                                if self.update_status:
                                    self.update_status(self.dispatch_location,
                                                    f"Still {len(stops2)} open stops after submit")
                        else:
                            # Stops cleared after Submit - remove from audit queue if it was added
                            if route in self.open_routes_reported:
                                self.open_routes_reported.discard(route)
                                logging.info(f"âœ… Route {route} removed from audit queue (all stops closed)")
                            logging.info("âœ… All stops closed after submit")
                            if self.update_status:
                                self.update_status(self.dispatch_location, "All stops closed after submit")

                    # back to messages
                    for sel in (INBOX_TAB_XPATH, MSG_TAB_XPATH, REFRESH_BTN_CSS):
                        try:
                            if sel.startswith("span"):
                                driver.find_element(By.CSS_SELECTOR, sel).click()
                            else:
                                driver.find_element(By.XPATH, sel).click()
                        except:
                            pass
                    time.sleep(1)

                # normal exit inner loop
                break

            except NoSuchWindowException as e:
                logging.warning(f"âš ï¸ Auto-Clear browser window closed: {e}")
                if self.update_status:
                    self.update_status(self.dispatch_location, "Window closed â€“ restartingâ€¦")
                try:
                    driver.quit()
                except:
                    pass
                # throttle restart
                if self.stop_event.wait(5):
                    break
                continue
            except Exception as e:
                logging.exception(f"âŒ Auto-Clear monitor error ({type(e).__name__}): {e}, restarting in 5s")
                if self.update_status:
                    self.update_status(self.dispatch_location, f"Error ({type(e).__name__}) â€“ restartingâ€¦")
                try:
                    driver.quit()
                except:
                    pass
                # throttle restart
                if self.stop_event.wait(5):
                    break
                continue

        # final cleanup & callback
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass
        logging.info(f"ðŸ›‘ Auto-Clear monitor for {self.dispatch_location} stopped")
        if self.finish_callback:
            self.finish_callback(self.dispatch_location)

# â”€â”€â”€ RoutR Chrome Extension local server â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class _RoutRExtensionServer(HTTPServer):
    """HTTPServer that does not log client-disconnect errors (WinError 10053/10054, etc.)."""
    def handle_error(self, request, client_address):
        exc_type, exc_val, _ = sys.exc_info()
        if exc_type in (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            return  # client closed connection; no need to log or traceback
        super().handle_error(request, client_address)


class _RoutRExtensionHandler(BaseHTTPRequestHandler):
    app_ref = None  # set by CarrierMonitorApp when starting server

    def do_GET(self):
        if self.path.rstrip('/') != '/routr-extension':
            self.send_error(404)
            return
        app = _RoutRExtensionHandler.app_ref
        enabled = getattr(app, 'routr_extension_enabled', False)
        name = (getattr(app, 'agent_display_name', None) or '').strip()
        signature = f"-CXPC {name}" if name else ""
        version = getattr(app, 'routr_version', None)
        payload = {"enabled": enabled, "signature": signature}
        if version:
            payload["version"] = version
        body = json.dumps(payload).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(body))
        self.end_headers()
        try:
            self.wfile.write(body)
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError) as _:
            # Client closed connection before we finished (e.g. extension tab closed, timeout).
            # WinError 10053/10054 on Windows; EPIPE/ECONNRESET elsewhere. Safe to ignore.
            pass

    def log_message(self, format, *args):
        pass  # suppress server log spam

# â”€â”€â”€ GUI App â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class CarrierMonitorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # skip-list for Auto-Clear: dispatch_location â†’ set of route IDs
        self.skip_clear_routes = {}
        # â”€â”€â”€ WATCHDOG state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self._last_read_update  = {}   # dispatch_location â†’ datetime
        self._last_clear_update = {}   # dispatch_location â†’ datetime
        self._last_sof_update   = None # datetime
        # track per-location audit threads
        self.running_audit_monitors = {}
        # â”€â”€â”€ App metadata â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.routr_version = "1.0.6"

        # â”€â”€â”€ SoF+ per route color state (to detect and log color changes) â”€â”€â”€
        self.sof_route_states = {}
        # track currently-open audit warning popups (keyed by "loc_route")
        self._audit_popups = {}
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title(f"ROUTR V{self.routr_version}")
        import sys
        # figure out where PyInstaller unpacked our data files
        if getattr(sys, "frozen", False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)

        icon_path = os.path.join(base_path, "swicon.ico")
        try:
            # 'default=' is required on Windows for .ico files
            self.iconbitmap(default=icon_path)
        except Exception as e:
            logging.warning(f"Could not set window icon: {e}")

        self.geometry("600x600")
        self.resizable(False, False)
        self.configure(bg="black")  # Keep black background for compatibility
        
        # Make window transparent
        self.attributes("-alpha", 0.9)  # 90% opacity - slight transparency
        self.center_window()
        
        # Widget panel state
        self.widget_panel_visible = False
        self.widget_panel = None

        self.username = ""
        self.password = ""
        # Dispatcher display name, e.g. "Shawn W."
        self.agent_display_name = ""
        # Label on Auto-Clear tab showing which name will be used in messages
        self.clear_signature_label = None
        # Manual CXPC-name editor UI on Auto-Clear tab
        self.clear_name_frame = None
        self.clear_name_entry = None

        # RoutR Chrome Extension: on/off state and local server for browser extension
        self.routr_extension_enabled = True
        self._routr_extension_server = None
        self._routr_extension_thread = None
        self._start_routr_extension_server()

        # â”€â”€â”€ Auto-sort stop event â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.sort_stop_event = threading.Event()

        self.login_frame               = None
        self.dispatch_frame            = None
        self.running_monitors          = {}
        # â”€â”€â”€ Dropbox-tab tracking â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.dropbox_driver = None
        self.dropbox_handle = None
        # Dropbox refresh cycle management
        self.dropbox_refresh_active = False
        # Start dropbox refresh cycle if monitors are running
        self._start_dropbox_refresh_cycle()
        self.locations_container       = None
        self.running_clear_monitors    = {}
        self.clear_locations_container = None

        self.init_login_ui()
        
        # Create the widget button
        self.create_widget_button()

        # Hook the window close
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Toggle console with Shift+F12
        self.bind_all(
            "<Shift-F12>",
            lambda e: self.toggle_console()
        )
        
        # Enable Auto-Sort tab with Shift+F6
        self.bind_all("<Shift-F6>", self._check_shift_f6)
        
        # Enable Auto Move button with Shift+F7
        self.bind_all("<Shift-F7>", self._check_shift_f7)
        
        # Track key state for shift-f6 combination
        self._shift_f6_pressed = False
        # Track key state for shift-f7 combination
        self._shift_f7_pressed = False
        # kick off the watchdog check every minute
        self.after(60_000, self._check_watchdog)
        
        # Update widget content every 5 seconds
        self.after(5000, self._update_widget_periodically)

    def _format_signed_message(self, text: str) -> str:
        """
        Append the CXPC signature to any outbound courier message,
        unless it's already present or we don't have a dispatcher name.
        """
        base = (text or "").rstrip()
        name = (self.agent_display_name or "").strip()
        if not name:
            return base
        signature = f"-CXPC {name}"
        # Avoid double-signing if it's already there
        if base.endswith(signature) or signature in base:
            return base
        # Add a space before the signature if needed
        sep = "" if base.endswith((".", "!", "?")) else "."
        return f"{base}{sep} {signature}"

    def _update_clear_signature_label(self):
        """
        Refresh the 'Sending clear message as ...' text and, when appropriate,
        mirror the current dispatcher name into the CXPC Name box.
        """
        try:
            if self.clear_signature_label is None:
                return
            if self.agent_display_name:
                txt = f"Sending clear message as -CXPC {self.agent_display_name}"
            else:
                txt = "Sending clear message without CXPC name (login or save a CXPC Name)."
            self.clear_signature_label.configure(text=txt)
            # If the manual box is empty and we have a name, prefill it
            if (
                self.clear_name_entry is not None
                and self.agent_display_name
                and not self.clear_name_entry.get().strip()
            ):
                self.clear_name_entry.delete(0, "end")
                self.clear_name_entry.insert(0, self.agent_display_name)
        except Exception as e:
            logging.warning(f"âš ï¸ Failed to refresh clear-signature label: {e!r}")

    def _set_agent_display_name_from_full(self, full_name: str):
        """
        Given a EmployeeDirectory full name like "Shawn Williams",
        store a compact display form such as "Shawn W." for CXPC signatures.
        """
        try:
            name = full_name.replace('\u00A0', ' ').strip()
            if not name:
                return
            parts = name.split()
            if len(parts) == 1:
                display = parts[0].title()
            else:
                first = parts[0].title()
                last_initial = parts[-1][0].upper()
                display = f"{first} {last_initial}."
            self.agent_display_name = display
            logging.info(f"âœ… Dispatcher display name set to {self.agent_display_name!r}")
            # Update the Auto-Clear signature label on the Tk main thread
            import threading
            if threading.current_thread() is threading.main_thread():
                self._update_clear_signature_label()
            else:
                try:
                    self.after(0, self._update_clear_signature_label)
                except Exception as e:
                    logging.warning(f"âš ï¸ Failed to schedule clear-signature label update: {e!r}")
        except Exception as e:
            logging.warning(f"âš ï¸ Failed to parse dispatcher display name from {full_name!r}: {e!r}")

    def _bind_sof_scroll_wheel_simple(self):
        """Simple scroll wheel binding for SoF+ grid frame"""
        def on_mousewheel(event):
            try:
                # Get the delta from the mouse wheel event
                delta = event.delta
                
                # Access the internal canvas of the scrollable frame
                if hasattr(self, 'sof_grid_frame') and hasattr(self.sof_grid_frame, '_parent_canvas'):
                    canvas = self.sof_grid_frame._parent_canvas
                    
                    # Calculate scroll amount for uniform, fast scrolling
                    # Use a larger multiplier for faster scrolling
                    scroll_amount = -1 * (delta // 120) * 50  # 50x faster scrolling
                    
                    # Perform the scroll
                    canvas.yview_scroll(scroll_amount, "units")
                    
                    # Prevent the event from propagating to other widgets
                    return "break"
            except Exception as e:
                # If anything goes wrong, just ignore the scroll event
                pass
        
        # Bind the mouse wheel event to the SoF+ grid frame
        try:
            self.sof_grid_frame.bind("<MouseWheel>", on_mousewheel)
        except:
            pass
        
        # Also bind to the internal canvas for better coverage
        try:
            if hasattr(self.sof_grid_frame, '_parent_canvas'):
                self.sof_grid_frame._parent_canvas.bind("<MouseWheel>", on_mousewheel)
        except:
            pass



    def safe_route_sort_key(self, route_key):
        """Safely convert route key to a sortable value, handling 'nan' and other edge cases."""
        try:
            # Handle 'nan' values
            if str(route_key).lower() == 'nan':
                return float('inf')  # Put 'nan' at the end
            
            # Try to convert to int, stripping leading zeros
            cleaned = str(route_key).lstrip('0') or '0'
            return int(cleaned)
        except (ValueError, TypeError):
            # If conversion fails, return a high value to put it at the end
            return float('inf')

    def center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
    
    def create_widget_button(self):
        """Create a smaller orange triangle widget button that can be repositioned."""
        # Create a smaller canvas for the triangle button
        self.widget_button_canvas = tk.Canvas(
            self, 
            width=40, 
            height=40, 
            bg="black", 
            highlightthickness=0,
            relief="flat"
        )
        # Fixed position in header (slightly inset from top-left of the black header band)
        self.widget_button_x = -5
        self.widget_button_y = -7
        # Initially hide the button - will show only on SOF+ tab
        self.widget_button_canvas.place_forget()
        
        # Draw smaller orange triangle
        triangle_points = [5, 5, 5, 35, 35, 5]  # Smaller triangle
        self.triangle_id = self.widget_button_canvas.create_polygon(
            triangle_points,
            fill=BRAND_ORANGE,
            outline=BRAND_ORANGE,
            width=2
        )
        
        # Bind click event (collapse to widget)
        self.widget_button_canvas.bind("<Button-1>", self.toggle_widget_panel)
        
        # Add hover effect
        self.widget_button_canvas.bind("<Enter>", self.on_widget_button_enter)
        self.widget_button_canvas.bind("<Leave>", self.on_widget_button_leave)
    
    def _start_move_widget_button(self, event):
        """Drag is disabled; widget button position is locked."""
        return "break"
    
    def _drag_widget_button(self, event):
        """Drag is disabled; widget button position is locked."""
        return "break"
    
    def _end_move_widget_button(self, event):
        """Drag is disabled; widget button position is locked."""
        return "break"
    
    def update_widget_button_visibility(self):
        """Show widget button only when on SOF+ tab"""
        # Cancel any pending updates to avoid callback errors
        if hasattr(self, '_widget_visibility_after_id'):
            try:
                self.after_cancel(self._widget_visibility_after_id)
            except (tk.TclError, ValueError):
                pass  # Callback may have already executed
        
        # Defer update slightly to avoid interfering with tab switch animation
        # Small delay allows the tab switch animation to complete smoothly
        try:
            self._widget_visibility_after_id = self.after(50, self._update_widget_button_visibility_deferred)
        except (tk.TclError, AttributeError):
            pass  # Window might be destroyed
    
    def _update_widget_button_visibility_deferred(self):
        """Actually update widget button visibility (deferred for smooth tab switching)"""
        # Clear the after_id since this callback is executing
        if hasattr(self, '_widget_visibility_after_id'):
            self._widget_visibility_after_id = None
            
        # Check if widgets still exist
        if not hasattr(self, 'tab'):
            return
        
        # Check if widgets are still valid
        try:
            if not self.tab.winfo_exists():
                return
        except (tk.TclError, AttributeError):
            return
            
        try:
            # Get current tab - this can fail if tabview is being destroyed
            try:
                current_tab = self.tab.get()
            except (tk.TclError, AttributeError) as e:
                # Tabview might be destroyed or not ready yet
                return
            
            if current_tab == "SoF+":
                # Always show widget button when on SoF+ tab.
                # Use a single overlay canvas attached to the main window so
                # it can sit up in the black header area.
                try:
                    if not hasattr(self, 'widget_button_canvas') or not self.widget_button_canvas.winfo_exists():
                        self.create_widget_button()
                    # Place using last remembered coordinates (default top-left)
                    x = getattr(self, "widget_button_x", 0)
                    y = getattr(self, "widget_button_y", 0)
                    self.widget_button_canvas.place(x=x, y=y, anchor="nw")
                    self.widget_button_canvas.lift()  # Bring to front to ensure visibility
                    self.update_idletasks()
                except (tk.TclError, AttributeError):
                    pass
            else:
                # Hide widget button when not on SoF+ tab
                if hasattr(self, 'widget_button_canvas'):
                    try:
                        if self.widget_button_canvas.winfo_exists():
                            # Only hide if it's an overlay (child of main window)
                            # If it's child of SoF+ tab, it will be hidden automatically
                            if self.widget_button_canvas.master == self:
                                self.widget_button_canvas.place_forget()
                    except (tk.TclError, AttributeError):
                        pass
        except (tk.TclError, AttributeError):
            # Ignore errors if widgets are being destroyed/recreated
            pass
        except Exception as e:
            # Log unexpected errors but don't spam - only log once per error type
            if not hasattr(self, '_last_widget_error') or self._last_widget_error != str(e):
                logging.debug(f"Unexpected error in widget button visibility: {e}")
                self._last_widget_error = str(e)
    
    def on_widget_button_enter(self, event):
        """Hover effect for widget button"""
        self.widget_button_canvas.itemconfig(self.triangle_id, fill="#FF8533")  # Lighter orange
    
    def on_widget_button_leave(self, event):
        """Reset widget button color on leave"""
        self.widget_button_canvas.itemconfig(self.triangle_id, fill=BRAND_ORANGE)
    
    def toggle_widget_panel(self, event=None):
        """Toggle the widget panel visibility"""
        if self.widget_panel_visible:
            self.hide_widget_panel()
        else:
            self.show_widget_panel()
    
    def show_widget_panel(self):
        """Show the draggable widget window and hide main program"""
        if self.widget_panel_visible:
            return
            
        # Hide the main program window
        self.withdraw()  # Hide the main window
        
        # Create a separate draggable widget window 
        self.widget_window = ctk.CTkToplevel()
        self.widget_window.title("ROUTR Widget")
        self.widget_window.geometry("560x240")  # Initial size, will be adjusted to fit content
        self.widget_window.configure(fg_color="black", bg="black")
        
        # Set the same icon as main program (before overrideredirect for best compatibility)
        icon_path = None
        try:
            if hasattr(sys, '_MEIPASS'):
                # Running as PyInstaller bundle
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(__file__)
            # Try routr.ico first, fallback to swicon.ico
            icon_path = os.path.join(base_path, "routr.ico")
            if not os.path.exists(icon_path):
                icon_path = os.path.join(base_path, "swicon.ico")
            if os.path.exists(icon_path):
                # Try multiple methods to set the icon on CTkToplevel
                # Method 1: wm_iconbitmap (standard tkinter method)
                try:
                    self.widget_window.wm_iconbitmap(icon_path)
                except:
                    pass
                # Method 2: iconbitmap with default parameter
                try:
                    self.widget_window.iconbitmap(default=icon_path)
                except:
                    pass
        except Exception as e:
            logging.warning(f"Could not set widget icon: {e}")
        
        # Remove window decorations AFTER setting icon
        self.widget_window.overrideredirect(True)
        
        # Set icon again after overrideredirect (sometimes needed for taskbar icon)
        # Try multiple times with delays as overrideredirect can interfere
        if icon_path and os.path.exists(icon_path):
            def set_icon_delayed(path, delay):
                try:
                    self.widget_window.wm_iconbitmap(path)
                except:
                    pass
            
            # Try immediately
            try:
                self.widget_window.wm_iconbitmap(icon_path)
            except:
                pass
            
            # Try after short delays
            self.widget_window.after(50, lambda: set_icon_delayed(icon_path, 50))
            self.widget_window.after(200, lambda: set_icon_delayed(icon_path, 200))
            self.widget_window.after(500, lambda: set_icon_delayed(icon_path, 500))
        
        # Modern styling - add transparency like main program
        self.widget_window.attributes("-alpha", 0.9)  # Match main program transparency
        self.widget_window.attributes("-topmost", True)  # Keep on top
        
        # Make it draggable
        self.make_draggable(self.widget_window)
        
        # Snap to screen edge
        self.snap_to_edge(self.widget_window)
        
        # Set up window close handler
        self.widget_window.protocol("WM_DELETE_WINDOW", self.return_to_main_program)
        
        self.widget_panel_visible = True
        
        # Create content for the widget panel (defer heavy operations)
        self.create_widget_content()
        
        # Defer periodic updates to avoid blocking window creation
        self.after(100, self._update_widget_periodically)
    
    def hide_widget_panel(self):
        """Hide the widget panel and return to main program"""
        if not self.widget_panel_visible:
            return
        
        # Set flag first to prevent race conditions
        self.widget_panel_visible = False
        
        # Show the main program window immediately (before destroying widget)
        self.deiconify()  # Show the main window
        self.lift()  # Bring to front
        
        # Destroy widget window asynchronously to avoid blocking
        if hasattr(self, 'widget_window') and self.widget_window:
            widget_to_destroy = self.widget_window
            self.widget_window = None
            try:
                # Destroy immediately but catch any errors
                widget_to_destroy.destroy()
            except:
                pass  # Widget window may already be destroyed
        
        # Focus handling (defer to avoid blocking)
        self.after(10, lambda: self.focus_force() if not self.widget_panel_visible else None)
        self.attributes('-topmost', True)  # Temporarily make topmost
        self.after(100, lambda: self.attributes('-topmost', False))  # Remove topmost after 100ms
    
    def return_to_main_program(self):
        """Return to main program from widget"""
        print("ðŸ”™ Attempting to return to main program...")
        try:
            self.hide_widget_panel()
            print("âœ… Successfully returned to main program")
        except Exception as e:
            print(f"âŒ Error returning to main program: {e}")
            # Force show main window as fallback
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except:
                pass
    
    def create_widget_content(self):
        """Create modern widget content with prominent restore buttons"""
        if not hasattr(self, 'widget_window'):
            return
            
        # Main frame with modern styling
        main_frame = ctk.CTkFrame(
            self.widget_window, 
            fg_color="black",
            corner_radius=6,
            border_width=1,
            border_color=BRAND_ORANGE
        )
        main_frame.pack(fill="x", padx=2, pady=2)  # Fill horizontally, let content determine height
        
        # Corner triangle restore button (top-left)
        restore_canvas = tk.Canvas(
            self.widget_window, 
            width=35, 
            height=35, 
            bg="black", 
            highlightthickness=0,
            relief="flat"
        )
        restore_canvas.place(x=5, y=5)
        
        # Draw triangle button
        triangle_points = [5, 5, 5, 30, 30, 5]
        self.widget_triangle_id = restore_canvas.create_polygon(
            triangle_points,
            fill=BRAND_ORANGE,
            outline=BRAND_ORANGE,
            width=2
        )
        
        # Bind click and hover events
        restore_canvas.bind("<Button-1>", lambda e: self.return_to_main_program())
        restore_canvas.bind("<Enter>", lambda e: restore_canvas.itemconfig(self.widget_triangle_id, fill="#FF8533"))
        restore_canvas.bind("<Leave>", lambda e: restore_canvas.itemconfig(self.widget_triangle_id, fill=BRAND_ORANGE))
        
        # Location display at the top
        location_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        location_frame.pack(fill="x", padx=5, pady=(10,5))
        
        self.location_label = ctk.CTkLabel(
            location_frame,
            text="Not Set",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        )
        self.location_label.pack()
        
        # Counters section
        counters_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        counters_frame.pack(fill="x", padx=5, pady=5)
        
        # Create counter displays with even width distribution
        self.counter_frames = {}
        counter_names = ["UNA", "EARLY", "LATE", "UNTRANSMITTED", "DUPS", "DEX 03", "PUX 03"]
        
        # Single row of 7 counters to mirror SoF+ layout
        for col in range(len(counter_names)):
            counters_frame.grid_columnconfigure(col, weight=1, uniform="counter")
        
        for col, name in enumerate(counter_names):
            row = 0  # all counters on a single row
            counter_frame = ctk.CTkFrame(
                counters_frame,
                fg_color="transparent",  # Match SOF+ style (no grey box)
                width=100,
                height=70,
                corner_radius=0,
                border_width=0,
                border_color="black"
            )
            counter_frame.grid(row=row, column=col, padx=2, pady=2, sticky="ew")
            counter_frame.pack_propagate(False)
            
            # Define click handler based on counter name
            def make_click_handler(counter_name):
                if counter_name == "UNA":
                    return lambda e: self.show_una_popup()
                elif counter_name == "EARLY":
                    return lambda e: self.show_early_popup()
                elif counter_name == "LATE":
                    return lambda e: self.show_late_popup()
                elif counter_name == "UNTRANSMITTED":
                    return lambda e: self.show_untransmitted_popup()
                elif counter_name == "DUPS":
                    return lambda e: self.show_duplicates_popup()
                elif counter_name == "DEX 03":
                    return lambda e: getattr(self, "show_dex03_popup", lambda: None)()
                elif counter_name == "PUX 03":
                    return lambda e: getattr(self, "show_pux03_popup", lambda: None)()
                return lambda e: None
            
            click_handler = make_click_handler(name)
            
            # Hover handlers that update the counter frame border
            hover_enter = lambda e, cf=counter_frame: cf.configure(border_color=BRAND_ORANGE)
            hover_leave = lambda e, cf=counter_frame: cf.configure(border_color="#000000")
            
            # Bind to counter frame
            counter_frame.bind("<Button-1>", click_handler)
            counter_frame.bind("<Enter>", hover_enter)
            counter_frame.bind("<Leave>", hover_leave)
            
            # Counter value - use AnimatedSpinnerCounter for spinner support (2/3 size = 24)
            # Create a frame to hold the counter widget (transparent look)
            counter_widget_frame = tk.Frame(counter_frame, bg="black")
            counter_widget_frame.pack(expand=True, fill="both", pady=(5, 0))
            
            # Bind events to counter_widget_frame
            counter_widget_frame.bind("<Button-1>", click_handler)
            counter_widget_frame.bind("<Enter>", hover_enter)
            counter_widget_frame.bind("<Leave>", hover_leave)
            
            # Create spinner counter with slightly smaller numbers than SOF+
            widget_counter = AnimatedSpinnerCounter(counter_widget_frame, size=26, color=BRAND_ORANGE)
            # Set backgrounds to blend with popup (no grey box)
            widget_counter.frame.configure(bg="black")
            widget_counter.text_label.configure(bg="black")
            widget_counter.canvas.configure(bg="black")
            widget_counter.show_text("0")
            widget_counter.pack(expand=True)
            
            # Bind events to all counter widget components
            widget_counter.frame.bind("<Button-1>", click_handler)
            widget_counter.frame.bind("<Enter>", hover_enter)
            widget_counter.frame.bind("<Leave>", hover_leave)
            widget_counter.canvas.bind("<Button-1>", click_handler)
            widget_counter.canvas.bind("<Enter>", hover_enter)
            widget_counter.canvas.bind("<Leave>", hover_leave)
            if hasattr(widget_counter, 'text_label') and widget_counter.text_label:
                widget_counter.text_label.bind("<Button-1>", click_handler)
                widget_counter.text_label.bind("<Enter>", hover_enter)
                widget_counter.text_label.bind("<Leave>", hover_leave)
            
            # Store reference to the counter widget
            self.counter_frames[name] = widget_counter
            
            # Counter label (match SOF+ style)
            name_label = ctk.CTkLabel(
                counter_frame,
                text="UNT" if name == "UNTRANSMITTED" else ("DUPS" if name == "DUPS" else name),
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color="#FF6600"
            )
            name_label.pack(pady=(0, 5))
            
            # Bind events + tooltips to name label
            name_label.bind("<Button-1>", click_handler)
            name_label.bind("<Enter>", hover_enter)
            name_label.bind("<Leave>", hover_leave)
            
            # Attach descriptive tooltips for UNA/EARLY/LATE/UNT/DUPS/DEX 03/PUX 03
            try:
                if name == "UNA":
                    CustomTooltip(name_label, "UNASSIGNED STOPS")
                elif name == "EARLY":
                    CustomTooltip(name_label, "EARLY PICK UPS")
                elif name == "LATE":
                    CustomTooltip(name_label, "LATE PICK UPS")
                elif name == "UNTRANSMITTED":
                    CustomTooltip(name_label, "UNTRANSMITTED STOPS")
                elif name == "DUPS":
                    CustomTooltip(name_label, "DUPLICATE STOPS")
                elif name == "DEX 03":
                    CustomTooltip(name_label, "BAD ADDRESSES")
                elif name == "PUX 03":
                    CustomTooltip(name_label, "BAD PICKUP ADDRESS")
            except Exception:
                pass
        
        # Problem routes section
        routes_header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        routes_header_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(
            routes_header_frame,
            text="âš ï¸ PROBLEM ROUTES",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#ff4444"
        ).pack()
        
        # Scrolling ticker frame for moving route boxes
        self.ticker_frame = ctk.CTkFrame(main_frame, fg_color="black", height=50, corner_radius=5, border_width=1, border_color=BRAND_ORANGE)
        self.ticker_frame.pack(fill="x", padx=5, pady=(0, 5))
        self.ticker_frame.pack_propagate(False)
        
        # Create canvas for scrolling route boxes - ensure it's visible and properly sized
        self.ticker_canvas = tk.Canvas(
            self.ticker_frame, 
            bg="black", 
            height=40, 
            width=540,  # Explicit width to ensure visibility
            highlightthickness=0,
            scrollregion=(0, 0, 2000, 40)  # Wide scrolling area
        )
        self.ticker_canvas.pack(fill="both", expand=True, padx=5, pady=5)
        # Ensure canvas is visible
        self.ticker_canvas.update_idletasks()
        
        # Initialize ticker variables
        self.ticker_boxes = []
        self.ticker_box_pairs = []  # Initialize ticker_box_pairs
        self.ticker_routes_set = set()  # Track routes in ticker
        self.ticker_position = 0
        self.ticker_width = 560  # Approximate widget width
        self.ticker_speed = 2  # Pixels per frame
        
        # Show initial "No Problem Routes" message if no routes exist
        if not hasattr(self, 'sof_route_states') or not self.sof_route_states:
            self.ticker_canvas.create_text(
                270, 20, 
                text="âœ… No Problem Routes", 
                fill="#00cc00", 
                font=("Arial", 16, "bold"),
                tags="no_routes_msg"
            )
        
        # Initialize route tracking
        if not hasattr(self, 'sof_route_states'):
            self.sof_route_states = {}
        self.all_problem_routes = []  # Store all routes for the popup
        
        # Update window size to fit content exactly (after all widgets are packed)
        def resize_to_content():
            self.widget_window.update_idletasks()
            main_frame.update_idletasks()
            # Get the actual height needed by the main frame plus padding
            content_height = main_frame.winfo_reqheight() + 4  # Add 4px for window padding
            # Set to exact content size
            widget_width = 560
            widget_height = content_height
            self.widget_window.geometry(f"{widget_width}x{widget_height}")
            # Update snap functions with new height
            if hasattr(self, 'widget_window'):
                self.widget_window.update_idletasks()
        
        # Resize after a short delay to ensure all widgets are rendered
        self.after(150, resize_to_content)
        
        # Defer heavy content updates to avoid blocking window creation
        self.after(50, lambda: self.update_widget_content())
        
        # Start ticker animation (will recreate boxes if needed) - defer this too
        self.after(100, lambda: self.start_ticker_animation())
    
    def start_ticker_animation(self):
        """Start the scrolling ticker animation for problem route boxes"""
        if hasattr(self, 'ticker_canvas') and hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
            try:
                # Ensure ticker_box_pairs is initialized
                if not hasattr(self, 'ticker_box_pairs'):
                    self.ticker_box_pairs = []
                # If we have routes but no ticker boxes, recreate them using button states (same as SoF+)
                if not self.ticker_box_pairs and hasattr(self, 'sof_route_buttons') and self.sof_route_buttons:
                    problem_routes = []
                    warning_routes = []  # widget ignores warnings
                    for route, btn in self.sof_route_buttons.items():
                        try:
                            route_str = str(route)
                            fg_color = btn.cget("fg_color")
                            border_color = btn.cget("border_color")
                            border_width = btn.cget("border_width")
                            is_problem = (fg_color == "#FF0000" or (border_color == "#FF0000" and border_width > 0))
                            if is_problem:
                                problem_routes.append(route_str)
                        except Exception:
                            continue
                    if problem_routes:
                        print("ðŸ”„ Starting ticker animation - recreating boxes from button states")
                        self.create_ticker_boxes(problem_routes, warning_routes)
                
                # Always update ticker boxes, even if empty (for smooth "No Problem Routes" display)
                self.update_ticker_boxes()
                # Schedule next update in 50ms for smooth scrolling
                self.after(50, self.start_ticker_animation)
            except Exception as e:
                print(f"âŒ Error in ticker animation: {e}")
                import traceback
                traceback.print_exc()
                # Retry after a delay
                self.after(1000, self.start_ticker_animation)
    
    def update_ticker_boxes(self):
        """Update the scrolling ticker box positions"""
        if not hasattr(self, 'ticker_canvas'):
            return
        
            # Ensure ticker_box_pairs exists and is initialized
            if not hasattr(self, 'ticker_box_pairs'):
                self.ticker_box_pairs = []
            # If we have routes but no ticker boxes, try to recreate them using button states (same as SoF+)
            if hasattr(self, 'sof_route_buttons') and self.sof_route_buttons:
                problem_routes = []
                warning_routes = []  # widget ignores warnings
                for route, btn in self.sof_route_buttons.items():
                    try:
                        route_str = str(route)
                        fg_color = btn.cget("fg_color")
                        border_color = btn.cget("border_color")
                        border_width = btn.cget("border_width")
                        is_problem = (fg_color == "#FF0000" or (border_color == "#FF0000" and border_width > 0))
                        if is_problem:
                            problem_routes.append(route_str)
                    except Exception:
                        continue
                if problem_routes:
                    print("ðŸ”„ Recreating ticker boxes - ticker_box_pairs was missing (using button states)")
                    self.create_ticker_boxes(problem_routes, warning_routes)
                    return
            
        if not self.ticker_box_pairs:
            return
            
        ticker_speed = getattr(self, 'ticker_speed', 2)  # Default to 2 pixels if not set
        
        # Move all ticker box pairs to the left first
        for pair in self.ticker_box_pairs:
            try:
                box_id = pair['box_id']
                text_id = pair['text_id']
                
                # Move both box and text together
                self.ticker_canvas.move(box_id, -ticker_speed, 0)
                self.ticker_canvas.move(text_id, -ticker_speed, 0)
            except (tk.TclError, KeyError):
                continue
        
        # Now check for boxes that scrolled off and reposition them
        # Find the current rightmost position after all moves
        rightmost_x = -float('inf')
        for pair in self.ticker_box_pairs:
            try:
                coords = self.ticker_canvas.coords(pair['box_id'])
                if coords and len(coords) >= 2:
                    rightmost_x = max(rightmost_x, coords[2])  # coords[2] is right edge of rectangle
            except (tk.TclError, KeyError):
                continue
        
        # Reposition boxes that scrolled off the left edge
        for pair in self.ticker_box_pairs:
            try:
                box_id = pair['box_id']
                text_id = pair['text_id']
                
                # Get current position of the box
                coords = self.ticker_canvas.coords(box_id)
                if coords and len(coords) >= 2 and coords[0] < -60:  # Box has scrolled off left edge
                    # Reposition to the right side after the rightmost box
                    # Add spacing (70 pixels) after the rightmost position
                    new_x = rightmost_x + 70
                    current_x = coords[0]
                    offset = new_x - current_x
                    
                    # Move both box and text to the new position
                    self.ticker_canvas.move(box_id, offset, 0)
                    self.ticker_canvas.move(text_id, offset, 0)
                    
                    # Update the rightmost position for next box that needs repositioning
                    new_coords = self.ticker_canvas.coords(box_id)
                    if new_coords and len(new_coords) >= 2:
                        rightmost_x = max(rightmost_x, new_coords[2])
                    
            except (tk.TclError, KeyError):
                # Box may have been destroyed
                continue
    
    def add_route_to_ticker(self, route, route_type="critical"):
        """Legacy incremental ticker API â€“ now a no-op to avoid duplicates/color drift."""
        return
    def remove_route_from_ticker(self, route):
        """Remove a route box dynamically from the ticker"""
        if not hasattr(self, 'ticker_canvas') or not hasattr(self, 'ticker_box_pairs'):
            return
        
        # Normalize route to string for consistent comparison
        route_str = str(route)
        
        # Find and remove all boxes for this route
        boxes_to_remove = []
        for i, pair in enumerate(self.ticker_box_pairs):
            if str(pair.get('route', '')) == route_str:
                boxes_to_remove.append(i)
        
        # Remove boxes in reverse order to maintain indices
        for i in reversed(boxes_to_remove):
            pair = self.ticker_box_pairs[i]
            try:
                self.ticker_canvas.delete(pair['box_id'])
                self.ticker_canvas.delete(pair['text_id'])
            except tk.TclError:
                pass
            self.ticker_box_pairs.pop(i)
        
        # Also remove from ticker_boxes
        if hasattr(self, 'ticker_boxes'):
            self.ticker_boxes = [box_id for pair in self.ticker_box_pairs 
                                for box_id in [pair['box_id'], pair['text_id']]]
        
        # Remove from tracking set
        if hasattr(self, 'ticker_routes_set'):
            self.ticker_routes_set.discard(route_str)
        
        print(f"âž– Removed route {route_str} from ticker dynamically")
        
        # If no routes left, show "No Problem Routes" message
        if not self.ticker_box_pairs and hasattr(self, 'ticker_canvas'):
            text_id = self.ticker_canvas.create_text(
                250, 20, 
                text="âœ… No Problem Routes", 
                fill="#00cc00", 
                font=("Arial", 12, "bold")
            )
            if hasattr(self, 'ticker_boxes'):
                self.ticker_boxes.append(text_id)
    
    def create_ticker_boxes(self, problem_routes, warning_routes):
        """Create scrolling route boxes on the ticker canvas (used for initial creation)"""
        if not hasattr(self, 'ticker_canvas'):
            print("âŒ Ticker canvas not found - cannot create ticker boxes")
            return
        
        try:
            # Ensure canvas is visible and updated
            self.ticker_canvas.update_idletasks()
            
            # Clear existing boxes
            self.ticker_canvas.delete("all")
            self.ticker_boxes = []  # Keep for backward compatibility
            self.ticker_box_pairs = []  # Store pairs of (box_id, text_id) for easier tracking
            self.ticker_routes_set = set()  # Track which routes are in the ticker
            
            # Build list of routes for ticker, ensuring no duplicates
            # (each route should appear only once in the widget)
            all_routes = list(dict.fromkeys(problem_routes + warning_routes))
            if not all_routes:
                # Show "No Problem Routes" message centered
                canvas_width = self.ticker_canvas.winfo_width() if self.ticker_canvas.winfo_width() > 1 else 540
                text_id = self.ticker_canvas.create_text(
                    canvas_width // 2, 20, 
                    text="âœ… No Problem Routes", 
                    fill="#00cc00", 
                    font=("Arial", 12, "bold")
                )
                self.ticker_boxes.append(text_id)
                print("âœ… Displaying 'No Problem Routes' message")
                return
            
            print(f"ðŸŽ¯ Creating ticker with {len(problem_routes)} problem routes for widget ticker")
            print(f"ðŸŽ¯ Problem routes: {problem_routes}")
            
            # Create a single set of boxes so routes are not visually repeated
            # (scroll logic will recycle this single set for continuous movement)
            box_width = 50
            box_spacing = 70
            canvas_width = self.ticker_canvas.winfo_width() if self.ticker_canvas.winfo_width() > 1 else 540
            set_width = len(all_routes) * box_spacing if all_routes else box_spacing
            num_sets = 1  # Only one logical set of routes; no duplicates
            
            # Update scrollregion to accommodate all boxes
            total_width = num_sets * set_width if set_width > 0 else canvas_width * 3
            self.ticker_canvas.config(scrollregion=(0, 0, total_width, 40))
            
            # Create route boxes spaced across the ticker - start from visible area
            x_position = 10  # Start slightly from left edge to be visible immediately
            
            for set_num in range(num_sets):
                for i, route in enumerate(all_routes):
                    # Normalize route to string for consistent comparison
                    route_str = str(route)
                    
                    # Determine box color by mirroring the actual SoF+ button colors
                    box_color = "#cc0000"
                    border_color = "#ff3333"
                    route_type = "critical"
                    try:
                        if hasattr(self, "sof_route_buttons") and route_str in self.sof_route_buttons:
                            btn = self.sof_route_buttons[route_str]
                            fg = btn.cget("fg_color")
                            bc = btn.cget("border_color")
                            # Fallbacks
                            if not fg:
                                fg = box_color
                            if not bc:
                                bc = fg
                            box_color = fg
                            border_color = bc
                    except Exception:
                        # Fall back to default critical colors if anything goes wrong
                        pass
                    
                    # Create box background
                    box_id = self.ticker_canvas.create_rectangle(
                        x_position, 8, x_position + box_width, 32,
                        fill=box_color, outline=border_color, width=2
                    )
                    
                    # Create route text
                    text_id = self.ticker_canvas.create_text(
                        x_position + 25, 20,
                        text=route_str, 
                        fill="#FFFFFF", 
                        font=("Arial", 10, "bold")
                    )
                    
                    # Store box pair together
                    self.ticker_box_pairs.append({
                        'box_id': box_id,
                        'text_id': text_id,
                        'route': route_str,
                        'route_type': route_type.lower()
                    })
                    
                    # Also store in ticker_boxes for backward compatibility
                    self.ticker_boxes.extend([box_id, text_id])
                    
                    # Track route in set (use normalized string)
                    self.ticker_routes_set.add(route_str)
                    
                    # Space boxes 70 pixels apart
                    x_position += box_spacing
            
            # Ensure canvas is updated and visible
            self.ticker_canvas.update_idletasks()
            print(f"âœ… Created {len(self.ticker_box_pairs)} ticker boxes")
        except Exception as e:
            print(f"âŒ Error creating ticker boxes: {e}")
            import traceback
            traceback.print_exc()
    
    def toggle_expandable_details(self):
        """Toggle the expandable details section for problem routes"""
        if not hasattr(self, 'details_frame'):
            return
            
        if self.details_expanded:
            # Hide the details
            self.details_frame.pack_forget()
            self.details_expanded = False
            # Update button text to show it's collapsed
            total_count = len(getattr(self, 'all_problem_routes', []))
            if total_count == 0:
                self.expand_button.configure(text="âœ… No Problem Routes")
            else:
                self.expand_button.configure(text=f"ðŸ“‹ Show All Problem Routes ({total_count}) â–¼")
        else:
            # Show the details
            self.update_expandable_details()
            self.details_frame.pack(fill="x", padx=5, pady=(0, 5))
            self.details_expanded = True
            # Update button text to show it's expanded
            total_count = len(getattr(self, 'all_problem_routes', []))
            self.expand_button.configure(text=f"ðŸ“‹ Hide Problem Routes ({total_count}) â–²")

    def update_expandable_details(self):
        """Update the content of the expandable details section"""
        if not hasattr(self, 'details_frame'):
            print("âŒ Details frame not found")
            return
        
        try:
            # Clear existing content
            for widget in self.details_frame.winfo_children():
                widget.destroy()
            
            if not hasattr(self, 'all_problem_routes') or not self.all_problem_routes:
                # No routes to show
                no_routes_label = ctk.CTkLabel(
                    self.details_frame,
                    text="âœ… No problem routes currently",
                    font=ctk.CTkFont(size=12),
                    text_color="#00cc00"
                )
                no_routes_label.pack(pady=10)
                return
            
            # Create scrollable area for routes with proper height
            routes_scroll = ctk.CTkScrollableFrame(
                self.details_frame,
                fg_color="black",
                height=150,  # Increased height for better visibility
                width=540
            )
            routes_scroll.pack(fill="both", expand=True, padx=5, pady=5)
            
            # Add routes grouped by type
            critical_routes = [route for route, route_type in self.all_problem_routes if route_type == "critical"]
            warning_routes = [route for route, route_type in self.all_problem_routes if route_type == "warning"]
            
            if critical_routes:
                # Critical routes header
                critical_header = ctk.CTkLabel(
                    routes_scroll,
                    text="ðŸš¨ Critical Routes:",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#ff3333"
                )
                critical_header.pack(anchor="w", pady=(5,2))
                
                # Critical routes list
                critical_text = ", ".join(str(route) for route in sorted(critical_routes, key=lambda x: int(x) if str(x).isdigit() else 999))
                critical_routes_label = ctk.CTkLabel(
                    routes_scroll,
                    text=critical_text,
                    font=ctk.CTkFont(size=11),
                    text_color="#ff6666",
                    wraplength=500
                )
                critical_routes_label.pack(anchor="w", padx=(10,5), pady=(0,5))
            
            if warning_routes:
                # Warning routes header  
                warning_header = ctk.CTkLabel(
                    routes_scroll,
                    text="âš ï¸ Warning Routes:",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#ffaa00"
                )
                warning_header.pack(anchor="w", pady=(5,2))
                
                # Warning routes list
                warning_text = ", ".join(str(route) for route in sorted(warning_routes, key=lambda x: int(x) if str(x).isdigit() else 999))
                warning_routes_label = ctk.CTkLabel(
                    routes_scroll,
                    text=warning_text,
                    font=ctk.CTkFont(size=11),
                    text_color="#ffcc66",
                    wraplength=500
                )
                warning_routes_label.pack(anchor="w", padx=(10,5), pady=(0,5))
        except Exception as e:
            print(f"âŒ Error updating expandable details: {e}")
            import traceback
            traceback.print_exc()

    def show_full_routes_popup(self):
        """Show a popup with all problem routes"""
        if not hasattr(self, 'all_problem_routes') or not self.all_problem_routes:
            return
            
        # Create popup window
        popup = ctk.CTkToplevel(self.widget_window)
        popup.title("All Problem Routes")
        popup.geometry("400x300")
        popup.configure(fg_color="#1a1a1a")
        popup.transient(self.widget_window)
        popup.grab_set()
        
        # Set icon
        try:
            if hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(__file__)
            icon_path = os.path.join(base_path, "swicon.ico")
            if os.path.exists(icon_path):
                popup.iconbitmap(default=icon_path)
        except Exception:
            pass
        
        # Title
        title_label = ctk.CTkLabel(
            popup,
            text="âš ï¸ All Problem Routes",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#ff4444"
        )
        title_label.pack(pady=10)
        
        # Scrollable frame for routes
        routes_frame = ctk.CTkScrollableFrame(
            popup, 
            fg_color="#2a2a2a",
            width=350,
            height=200
        )
        routes_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Add all routes to the popup
        for i, (route, route_type) in enumerate(self.all_problem_routes):
            route_button = ctk.CTkButton(
                routes_frame,
                text=f"Route {route}",
                font=ctk.CTkFont(size=12, weight="bold"),
                width=100,
                height=30,
                fg_color="#cc0000" if route_type == "critical" else "#ff6600",
                hover_color="#ff3333" if route_type == "critical" else "#ff9933",
                command=lambda r=route: self._route_clicked_popup(r, popup)
            )
            route_button.grid(row=i//3, column=i%3, padx=5, pady=5, sticky="ew")
            
        # Configure grid weights for responsive layout
        for col in range(3):
            routes_frame.grid_columnconfigure(col, weight=1)
        
        # Close button
        close_button = ctk.CTkButton(
            popup,
            text="Close",
            command=popup.destroy,
            fg_color="#666666",
            hover_color="#777777"
        )
        close_button.pack(pady=10)
        
        # Center the popup
        popup.update_idletasks()
        x = self.widget_window.winfo_x() + (self.widget_window.winfo_width() // 2) - (popup.winfo_width() // 2)
        y = self.widget_window.winfo_y() + (self.widget_window.winfo_height() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{x}+{y}")
    
    def _route_clicked_popup(self, route, popup):
        """Handle route click in popup - same functionality as main program"""
        popup.destroy()
        # Call the same route click handler as the main program would
        if hasattr(self, 'sof_route_states') and route in self.sof_route_states:
            self._get_route_status_info(route)
    
    def make_draggable(self, window):
        """Make a window draggable (no forced edge snapping so it can move across screens)."""
        def start_move(event):
            window.x = event.x
            window.y = event.y
        
        def stop_move(event):
            window.x = None
            window.y = None
        
        def do_move(event):
            if window.x is not None and window.y is not None:
                deltax = event.x - window.x
                deltay = event.y - window.y
                x = window.winfo_x() + deltax
                y = window.winfo_y() + deltay
                window.geometry(f"+{x}+{y}")
        
        window.bind("<Button-1>", start_move)
        window.bind("<ButtonRelease-1>", stop_move)
        window.bind("<B1-Motion>", do_move)
    
    def snap_to_edge(self, window):
        """Snap window to screen edge"""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        # Get actual window dimensions
        widget_width = window.winfo_width()
        widget_height = window.winfo_height()
        
        # Position at right edge of screen (compact widget size)
        x = screen_width - (widget_width + 20)  # Widget width + small margin
        y = 50  # Small offset from top
        window.geometry(f"{widget_width}x{widget_height}+{x}+{y}")
    
    def check_edge_proximity(self, window, x, y):
        """Check if window is near screen edges and snap if close"""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        widget_width = window.winfo_width()
        widget_height = window.winfo_height()
        
        snap_threshold = 20  # pixels from edge to trigger snap
        
        # Check right edge
        if x + widget_width >= screen_width - snap_threshold:
            x = screen_width - widget_width
            window.geometry(f"+{x}+{y}")
            return True
        
        # Check left edge
        if x <= snap_threshold:
            x = 0
            window.geometry(f"+{x}+{y}")
            return True
        
        # Check top edge
        if y <= snap_threshold:
            y = 0
            window.geometry(f"+{x}+{y}")
            return True
        
        # Check bottom edge
        if y + widget_height >= screen_height - snap_threshold:
            y = screen_height - widget_height
            window.geometry(f"+{x}+{y}")
            return True
        
        return False
    
    def snap_to_nearest_edge(self, window):
        """Snap window to the nearest screen edge"""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        widget_width = window.winfo_width()
        widget_height = window.winfo_height()
        
        current_x = window.winfo_x()
        current_y = window.winfo_y()
        
        # Calculate distances to each edge
        dist_to_left = current_x
        dist_to_right = screen_width - (current_x + widget_width)
        dist_to_top = current_y
        dist_to_bottom = screen_height - (current_y + widget_height)
        
        # Find the closest edge
        min_dist = min(dist_to_left, dist_to_right, dist_to_top, dist_to_bottom)
        
        if min_dist == dist_to_left:
            # Snap to left edge
            window.geometry(f"{widget_width}x{widget_height}+0+{current_y}")
        elif min_dist == dist_to_right:
            # Snap to right edge
            window.geometry(f"{widget_width}x{widget_height}+{screen_width - widget_width}+{current_y}")
        elif min_dist == dist_to_top:
            # Snap to top edge
            window.geometry(f"{widget_width}x{widget_height}+{current_x}+0")
        else:
            # Snap to bottom edge
            window.geometry(f"{widget_width}x{widget_height}+{current_x}+{screen_height - widget_height}")
    
    def hide_main_ui(self):
        """Hide all main UI elements"""
        # Store references to main UI elements
        self.main_ui_elements = []
        
        # Hide all children of the main window except the widget button
        for child in self.winfo_children():
            if child != self.widget_button_canvas:
                self.main_ui_elements.append(child)
                child.pack_forget()
    
    def show_main_ui(self):
        """Restore all main UI elements"""
        # Restore all main UI elements
        for element in getattr(self, 'main_ui_elements', []):
            try:
                element.pack()
            except:
                # If pack fails, try other geometry managers
                try:
                    element.place()
                except:
                    pass
    
    def update_widget_content(self):
        """Update the compact widget content"""
        if not self.widget_panel_visible or not hasattr(self, 'widget_window'):
            return
        
        # Update pickup counters with actual values from the main counters
        # Try to get values from the counter objects if they exist, otherwise use stored values
        counter_values = {}
        
        # Get UNA count
        if hasattr(self, 'una_counter') and hasattr(self.una_counter, 'current_count'):
            counter_values["UNA"] = self.una_counter.current_count
        else:
            counter_values["UNA"] = getattr(self, 'una_count', 0)
        
        # Get EARLY count  
        if hasattr(self, 'early_counter') and hasattr(self.early_counter, 'current_count'):
            counter_values["EARLY"] = self.early_counter.current_count
        else:
            counter_values["EARLY"] = getattr(self, 'early_count', 0)
            
        # Get LATE count
        if hasattr(self, 'late_counter') and hasattr(self.late_counter, 'current_count'):
            counter_values["LATE"] = self.late_counter.current_count
        else:
            counter_values["LATE"] = getattr(self, 'late_count', 0)
            
        # Get UNTRANSMITTED count
        if hasattr(self, 'untransmitted_counter') and hasattr(self.untransmitted_counter, 'current_count'):
            counter_values["UNTRANSMITTED"] = self.untransmitted_counter.current_count
        else:
            counter_values["UNTRANSMITTED"] = getattr(self, 'untransmitted_count', 0)
            
        # Get DUPLICATES count
        if hasattr(self, 'duplicates_counter') and hasattr(self.duplicates_counter, 'current_count'):
            counter_values["DUPS"] = self.duplicates_counter.current_count
        else:
            counter_values["DUPS"] = getattr(self, 'duplicates_count', 0)

        # Placeholder counts for DEX 03 / PUX 03 (widget only for now)
        counter_values["DEX 03"] = 0
        counter_values["PUX 03"] = 0
        
        # Update counter displays - check if main counters are spinning
        if hasattr(self, 'counter_frames'):
            counter_mapping = {
                "UNA": "una_counter",
                "EARLY": "early_counter",
                "LATE": "late_counter",
                "UNTRANSMITTED": "untransmitted_counter",
                "DUPS": "duplicates_counter",
                # DEX 03 / PUX 03 are not wired yet; keep them non-spinning placeholders
                "DEX 03": None,
                "PUX 03": None,
            }
            
            for name, value in counter_values.items():
                if name in self.counter_frames:
                    widget_counter = self.counter_frames[name]
                    # Check if main counter is spinning
                    main_counter_name = counter_mapping.get(name)
                    is_spinning = False
                    if main_counter_name and hasattr(self, main_counter_name):
                        main_counter = getattr(self, main_counter_name)
                        if hasattr(main_counter, 'is_spinning'):
                            is_spinning = main_counter.is_spinning
                    
                    # Show spinner or text based on main counter state
                    if is_spinning:
                        widget_counter.show_spinner()
                    else:
                        widget_counter.show_text(str(value))
                    print(f"ðŸ”„ Widget updated {name}: {value} (spinning: {is_spinning})")  # Debug output
        
        # Update location display from SOF+ location entry
        if hasattr(self, 'location_label'):
            if hasattr(self, 'sof_location_entry') and self.sof_location_entry.get().strip():
                location_text = f"{self.sof_location_entry.get().strip().upper()}"
            elif hasattr(self, 'running_monitors') and self.running_monitors:
                locations = list(self.running_monitors.keys())
                location_text = f"{', '.join(locations[:2])}"
                if len(locations) > 2:
                    location_text += f" +{len(locations)-2}"
            else:
                location_text = "Not Set"
            self.location_label.configure(text=location_text)
        
        # Update problem routes data and ticker
        # Initialize sof_route_states if it doesn't exist
        if not hasattr(self, 'sof_route_states'):
            self.sof_route_states = {}
            
        # Get problem routes (RED) only.
        # Use the SAME logic as SoF+ grid: check actual button states, not sof_route_states dict.
        # This ensures widget shows the same problem routes as SoF+ screen.
        problem_routes = []
        warning_routes = []  # kept for ticker API compatibility but we do NOT add to it
        
        # Check actual button states (same logic as SoF+ grid display)
        if hasattr(self, 'sof_route_buttons') and self.sof_route_buttons:
            print(f"ðŸ” Checking {len(self.sof_route_buttons)} route buttons for problems (matching SoF+ logic)...")
            for route, btn in self.sof_route_buttons.items():
                try:
                    route_str = str(route)
                    # Check if button has red border or red fill (same as SoF+ grid)
                    fg_color = btn.cget("fg_color")
                    border_color = btn.cget("border_color")
                    border_width = btn.cget("border_width")
                    
                    # Red fill (#FF0000) or red border with width > 0 = problem route
                    is_problem = (fg_color == "#FF0000" or 
                                (border_color == "#FF0000" and border_width > 0))
                    
                    if is_problem:
                        problem_routes.append(route_str)
                        print(f"ðŸš¨ Added route {route_str} to PROBLEM routes (from button state)")
                    # NOTE: Yellow/warning routes (on break) are intentionally
                    # ignored for the widget per user request.
                except Exception as e:
                    print(f"âš ï¸ Error checking route {route}: {e}")
                    continue
        
        # Combine and sort routes used for problem-route counts (only true problems)
        all_problem_routes = problem_routes
        all_problem_routes.sort(key=lambda x: int(x) if str(x).isdigit() else 999)
        
        # Store all routes with their types for the popup
        self.all_problem_routes = []
        for route in problem_routes:
            self.all_problem_routes.append((route, "critical"))
        
        # Dynamically update ticker boxes - add new routes and remove routes that are no longer problems
        if hasattr(self, 'ticker_canvas') and hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
            # Initialize ticker tracking if needed
            if not hasattr(self, 'ticker_routes_set'):
                self.ticker_routes_set = set()
            if not hasattr(self, 'ticker_box_pairs'):
                self.ticker_box_pairs = []
            
            # If ticker_box_pairs is empty but we have routes, recreate from scratch to ensure proper structure
            # Also check if canvas has "No Problem Routes" message that needs to be cleared
            has_no_routes_msg = False
            try:
                if hasattr(self, 'ticker_canvas'):
                    canvas_items = self.ticker_canvas.find_all()
                    for item in canvas_items:
                        try:
                            item_text = self.ticker_canvas.itemcget(item, "text")
                            if "No Problem Routes" in item_text:
                                has_no_routes_msg = True
                                break
                        except:
                            pass
            except:
                pass
            
            if (not self.ticker_box_pairs or has_no_routes_msg) and problem_routes:
                print(f"ðŸ”„ Ticker boxes empty or showing 'No Problem Routes' but {len(problem_routes)} routes exist - recreating ticker structure")
                # Normalize routes before passing to create_ticker_boxes
                problem_routes_normalized = [str(r) for r in problem_routes]
                warning_routes_normalized = [str(r) for r in warning_routes]
                self.create_ticker_boxes(problem_routes_normalized, warning_routes_normalized)
            else:
                # Normalize all routes to strings for consistent comparison
                problem_routes_str = [str(r) for r in problem_routes]
                warning_routes_str = [str(r) for r in warning_routes]
                
                # Get current routes in ticker (normalize to strings)
                current_ticker_routes = set(str(r) for r in (self.ticker_routes_set if hasattr(self, 'ticker_routes_set') else set()))
                
                # Create sets of new problem routes (already normalized)
                new_problem_routes = set(problem_routes_str)
                new_warning_routes = set()  # widget ignores warnings
                new_all_routes = new_problem_routes
                
                # Routes to add (in new routes but not in ticker)
                routes_to_add = new_all_routes - current_ticker_routes
                
                # Routes to remove (in ticker but not in new routes)
                routes_to_remove = current_ticker_routes - new_all_routes
                
                # Routes that changed type (need to update)
                routes_to_update = set()
                if hasattr(self, 'ticker_box_pairs') and self.ticker_box_pairs:
                    for pair in self.ticker_box_pairs:
                        route = str(pair.get('route', ''))
                        current_type = pair.get('route_type', 'critical')
                        if route in new_problem_routes and current_type != 'critical':
                            routes_to_update.add((route, 'critical'))
                        elif route in new_warning_routes and current_type != 'warning':
                            routes_to_update.add((route, 'warning'))
                
                # Remove routes that are no longer problems
                for route in routes_to_remove:
                    self.remove_route_from_ticker(route)
                
                # Update routes that changed type
                for route, new_type in routes_to_update:
                    self.remove_route_from_ticker(route)
                    self.add_route_to_ticker(route, new_type)
                
                # Add new problem routes dynamically
                for route in routes_to_add:
                    if route in new_problem_routes:
                        self.add_route_to_ticker(route, "critical")
                    elif route in new_warning_routes:
                        self.add_route_to_ticker(route, "warning")
                
                # If ticker is empty and we have no routes, ensure "No Problem Routes" message is shown
                if not new_all_routes and (not hasattr(self, 'ticker_box_pairs') or not self.ticker_box_pairs):
                    if hasattr(self, 'ticker_canvas'):
                        self.ticker_canvas.delete("all")
                        text_id = self.ticker_canvas.create_text(
                            250, 20, 
                            text="âœ… No Problem Routes", 
                            fill="#00cc00", 
                            font=("Arial", 12, "bold")
                        )
                        if hasattr(self, 'ticker_boxes'):
                            self.ticker_boxes = [text_id]
                        if hasattr(self, 'ticker_routes_set'):
                            self.ticker_routes_set = set()
        else:
            # Fallback: recreate ticker if widget panel not visible or canvas not ready
            self.create_ticker_boxes(problem_routes, warning_routes)
        
        # Update expand button text
        if hasattr(self, 'expand_button'):
            total_count = len(all_problem_routes)
            if total_count == 0:
                self.expand_button.configure(text="âœ… No Problem Routes", state="disabled")
                # Hide details if expanded and no routes
                if hasattr(self, 'details_expanded') and self.details_expanded:
                    self.details_frame.pack_forget()
                    self.details_expanded = False
            else:
                self.expand_button.configure(state="normal")
                # Update button text based on current state
                if hasattr(self, 'details_expanded') and self.details_expanded:
                    self.expand_button.configure(text=f"ðŸ“‹ Hide Problem Routes ({total_count}) â–²")
                    # Update the details content if currently expanded
                    self.update_expandable_details()
                else:
                    self.expand_button.configure(text=f"ðŸ“‹ Show All Problem Routes ({total_count}) â–¼")
    
    def _update_widget_periodically(self):
        """Periodically update widget content - only when data changes"""
        if self.widget_panel_visible:
            # Check if any data has actually changed before updating
            if self._widget_data_changed():
                print("ðŸ”„ Widget data changed, updating...")
                self.update_widget_content()
                # Store current state for next comparison
                self._store_widget_state()
            
        # Schedule next update in 5 seconds for responsive updates
        self.after(5000, self._update_widget_periodically)

    def _widget_data_changed(self):
        """Check if widget-relevant data has changed since last update"""
        # Initialize previous state if not exists
        if not hasattr(self, '_widget_prev_state'):
            return True  # First run, always update
            
        # Get current counter values
        current_counters = {}
        counter_attrs = ['una_count', 'early_count', 'late_count', 'untransmitted_count', 'duplicates_count']
        for attr in counter_attrs:
            if hasattr(self, f'{attr[:-6]}_counter') and hasattr(getattr(self, f'{attr[:-6]}_counter'), 'current_count'):
                current_counters[attr] = getattr(getattr(self, f'{attr[:-6]}_counter'), 'current_count')
            else:
                current_counters[attr] = getattr(self, attr, 0)
        
        # Get current sof+ route states
        current_routes = dict(getattr(self, 'sof_route_states', {}))
        
        # Get current location
        current_location = ""
        if hasattr(self, 'sof_location_entry') and self.sof_location_entry.get().strip():
            current_location = self.sof_location_entry.get().strip().upper()
        elif hasattr(self, 'running_monitors') and self.running_monitors:
            locations = list(self.running_monitors.keys())
            current_location = ', '.join(locations[:2])
            if len(locations) > 2:
                current_location += f" +{len(locations)-2}"
        
        # Compare with previous state
        prev_state = self._widget_prev_state
        
        # Check if counters changed
        if current_counters != prev_state.get('counters', {}):
            return True
            
        # Check if route states changed  
        if current_routes != prev_state.get('routes', {}):
            return True
            
        # Check if location changed
        if current_location != prev_state.get('location', ''):
            return True
            
        return False  # No changes detected

    def _store_widget_state(self):
        """Store current widget state for change detection"""
        # Get current counter values
        current_counters = {}
        counter_attrs = ['una_count', 'early_count', 'late_count', 'untransmitted_count', 'duplicates_count']
        for attr in counter_attrs:
            if hasattr(self, f'{attr[:-6]}_counter') and hasattr(getattr(self, f'{attr[:-6]}_counter'), 'current_count'):
                current_counters[attr] = getattr(getattr(self, f'{attr[:-6]}_counter'), 'current_count')
            else:
                current_counters[attr] = getattr(self, attr, 0)
        
        # Get current sof+ route states
        current_routes = dict(getattr(self, 'sof_route_states', {}))
        
        # Get current location
        current_location = ""
        if hasattr(self, 'sof_location_entry') and self.sof_location_entry.get().strip():
            current_location = self.sof_location_entry.get().strip().upper()
        elif hasattr(self, 'running_monitors') and self.running_monitors:
            locations = list(self.running_monitors.keys())
            current_location = ', '.join(locations[:2])
            if len(locations) > 2:
                current_location += f" +{len(locations)-2}"
        
        # Store current state
        self._widget_prev_state = {
            'counters': current_counters.copy(),
            'routes': current_routes.copy(),
            'location': current_location
        }

    def clear_frame(self, frame):
        for w in frame.winfo_children():
            w.destroy()
    
    def toggle_console(self):
        """Toggle console window visibility with Shift+F12 and capture all detailed logs"""
        try:
            import ctypes
            import logging
            
            if not hasattr(self, "console_allocated_general"):
                # First time - allocate console
                ctypes.windll.kernel32.AllocConsole()
                self.console_allocated_general = True
                self.console_visible = True
                
                # Create a custom console handler that writes directly to CONOUT$
                class ConsoleHandler(logging.Handler):
                    def __init__(self):
                        super().__init__()
                        self.console_output = open('CONOUT$', 'w', encoding='utf-8')
                    
                    def emit(self, record):
                        try:
                            msg = self.format(record)
                            self.console_output.write(msg + '\n')
                            self.console_output.flush()
                        except Exception:
                            pass
                    
                    def close(self):
                        if hasattr(self, 'console_output'):
                            self.console_output.close()
                
                # Create and configure the console handler
                console_handler = ConsoleHandler()
                console_handler.setLevel(logging.INFO)
                formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
                console_handler.setFormatter(formatter)
                
                # Get the root logger and add our console handler
                root_logger = logging.getLogger()
                root_logger.addHandler(console_handler)
                
                # Store the handler so we can remove it later
                self.console_handler = console_handler
                
                # Also capture print statements by redirecting sys.stdout
                import sys
                self.original_stdout = sys.stdout
                self.console_stdout = open('CONOUT$', 'w', encoding='utf-8')
                sys.stdout = self.console_stdout
                
                # Test the console
                print("ðŸ” Console opened with Shift+F12")
                print("ðŸ” All detailed logs will now appear here (like running from terminal)")
                print("ðŸ” Press Shift+F12 again to hide/show this console")
                print("=" * 60)
                
                # Show current application status
                logging.info("ðŸ” ROUTR Application Status:")
                logging.info(f"   - Active monitors: {len(self.running_monitors)}")
                logging.info(f"   - Clear monitors: {len(self.running_clear_monitors)}")
                logging.info(f"   - Auto-WalkUp enabled: {getattr(self, 'auto_walkup_var', None) and self.auto_walkup_var.get()}")
                
                # Show any recent activity
                if hasattr(self, '_last_read_update'):
                    for loc, last_update in self._last_read_update.items():
                        if last_update:
                            logging.info(f"   - {loc} last read: {last_update.strftime('%H:%M:%S')}")
                
                print("=" * 60)
                logging.info("âœ… Console logging activated - all logs will now appear here")
                
            else:
                # Console exists - toggle visibility
                console_window = ctypes.windll.kernel32.GetConsoleWindow()
                if console_window:
                    if hasattr(self, 'console_visible') and self.console_visible:
                        # Hide console
                        ctypes.windll.user32.ShowWindow(console_window, 0)  # SW_HIDE
                        self.console_visible = False
                        print("ðŸ” Console hidden")
                    else:
                        # Show console
                        ctypes.windll.user32.ShowWindow(console_window, 5)  # SW_SHOW
                        self.console_visible = True
                        print("ðŸ” Console shown")
                else:
                    print("ðŸ” Console window not found")
        except Exception as e:
            print(f"âŒ Error toggling console: {e}")
            # Reset flags on error
            if hasattr(self, "console_allocated_general"):
                self.console_allocated_general = False
            if hasattr(self, "console_visible"):
                self.console_visible = False
    
    def _check_shift_f6(self, event):
        """Check if shift-f6 is pressed to toggle Auto-Sort tab visibility"""
        try:
            # Initialize flag if not exists
            if not hasattr(self, '_shift_f6_pressed'):
                self._shift_f6_pressed = False
            
            # Check if Auto-Sort tab exists
            try:
                self.tab.tab("Auto-Sort")
                tab_exists = True
            except:
                tab_exists = False
            
            if tab_exists:
                # Tab exists - toggle visibility
                if self._shift_f6_pressed:
                    # Hide the Auto-Sort tab by removing it
                    self.tab.delete("Auto-Sort")
                    self._shift_f6_pressed = False
                    print("ðŸ” Auto-Sort tab hidden")
                else:
                    # Show the Auto-Sort tab by recreating it
                    self.tab.add("Auto-Sort")
                    sort_tab = self.tab.tab("Auto-Sort")
                    self._create_auto_sort_ui(sort_tab)
                    self._shift_f6_pressed = True
                    
                    # Display any stored results when showing the tab
                    if hasattr(self, 'stored_results') and self.stored_results:
                        print(f"ðŸ” Displaying {len(self.stored_results)} stored results...")
                        stored_copy = self.stored_results.copy()
                        self.stored_results.clear()
                        for result in stored_copy:
                            print(f"ðŸ” Adding stored result: {result[:50]}...")
                            self.add_result(result)
                    else:
                        print("ðŸ” No stored results to display")
            else:
                # Tab doesn't exist - create it for the first time
                print("ðŸ” Creating Auto-Sort tab for the first time...")
                self.tab.add("Auto-Sort")
                sort_tab = self.tab.tab("Auto-Sort")
                self._create_auto_sort_ui(sort_tab)
                
                # The tab is automatically shown when created, so just set the flag
                self._shift_f6_pressed = True
                
                # Display any stored results
                if hasattr(self, 'stored_results') and self.stored_results:
                    print(f"ðŸ” Displaying {len(self.stored_results)} stored results...")
                    stored_copy = self.stored_results.copy()
                    self.stored_results.clear()
                    for result in stored_copy:
                        print(f"ðŸ” Adding stored result: {result[:50]}...")
                        self.add_result(result)
                else:
                    print("ðŸ” No stored results to display")
                    
        except Exception as e:
            print(f"Error toggling Auto-Sort tab: {e}")
        
        # Return "break" to prevent event propagation to other widgets
        return "break"
    
    def start_auto_move(self):
        """Start the auto move process for selected results"""
        try:
            from tkinter import messagebox
            import re
            import threading
            
            # Check if user is logged in
            if not self.username or not self.password:
                messagebox.showerror("Not Logged In", "Please log in first before using auto move.")
                return

            # Check if we have mismatch data to process
            if not hasattr(self, 'result_frames') or not self.result_frames:
                messagebox.showerror("No Data", "Please run the route comparison first to get mismatch data.")
                return

            # Get mismatch data from results (only selected checkboxes)
            mismatch_data = []
            selected_count = 0
            total_mismatches = 0
            
            for result in self.result_frames:
                text = result['text']
                if text and text.strip().startswith("âŒ") and "Disp#" in text:
                    total_mismatches += 1
                    
                    # Check if this mismatch is selected via checkbox
                    if result['checkbox'] is not None and result['checkbox_var'].get():
                        selected_count += 1
                        
                        # Parse the mismatch data
                        # Format: "âŒ Disp#6192 | Route: 459 â†’ 779 | Customer: CLOSED SOUTHWEST FLORIDA LLC"
                        try:
                            # Extract dispatch number
                            disp_match = re.search(r'Disp#(\d+)', text)
                            if disp_match:
                                disp_no = disp_match.group(1)
                                
                                # Extract route information
                                route_match = re.search(r'Route: (\d+) â†’ (\d+)', text)
                                if route_match:
                                    current_route = route_match.group(1)
                                    correct_route = route_match.group(2)
                                    
                                    # Extract customer name
                                    customer_match = re.search(r'Customer: (.+)$', text)
                                    customer = customer_match.group(1) if customer_match else "Unknown"
                                    
                                    mismatch_data.append({
                                        'disp_no': disp_no,
                                        'current_route': current_route,
                                        'correct_route': correct_route,
                                        'customer': customer
                                    })
                        except Exception as e:
                            print(f"Error parsing mismatch data: {e}")
                            continue

            if not mismatch_data:
                messagebox.showerror("No Mismatches", f"No route mismatches selected to process.\n\nTotal mismatches found: {total_mismatches}\nSelected: {selected_count}")
                return

            # Confirm with user
            confirm_msg = f"Found {selected_count} selected route mismatches to fix (out of {total_mismatches} total).\n\n"
            confirm_msg += "This will automatically:\n"
            confirm_msg += "1. Open Pickup Request for each dispatch\n"
            confirm_msg += "2. Enter the location and dispatch number\n"
            confirm_msg += "3. Click Assign Route and enter the correct route\n"
            confirm_msg += "4. Click OK to save\n\n"
            confirm_msg += "Continue?"
            
            if not messagebox.askyesno("Confirm Auto Move", confirm_msg):
                return

            # Disable the Auto Move button during processing
            if hasattr(self, 'auto_move_button'):
                self.auto_move_button.configure(state="disabled")
            
            # Disable Start and Stop buttons to prevent accidental activation
            if hasattr(self, 'sort_start_button'):
                self.sort_start_button.configure(state="disabled")
            if hasattr(self, 'sort_stop_button'):
                self.sort_stop_button.configure(state="disabled")
            
            # Unbind Enter key from location entry to prevent accidental route mismatch activation
            if hasattr(self, 'sort_location_entry'):
                self.sort_location_entry.unbind("<Return>")
            
            # Clear all checkboxes when starting auto move (only if UI exists)
            if hasattr(self, 'select_all_var'):
                self.select_all_var.set(False)
            if hasattr(self, 'result_frames'):
                for result in self.result_frames:
                    if result['checkbox'] is not None:
                        result['checkbox_var'].set(False)
            
            # Hide checkboxes from GUI when auto move starts
            self.hide_checkboxes()
            print("ðŸ” Checkboxes hidden for auto move process")
            
            # Clear results and show status (don't show select all checkbox during auto move)
            self.clear_results(show_select_all=False)
            self.sort_status_label.configure(text="Processing Auto Move")
            
            # Start breathing animation for auto move
            if not hasattr(self, 'sort_breathing_alpha'):
                self.sort_breathing_alpha = 1.0
                self.sort_breathing_direction = -1
                self._breathing_animation()
            
            # Initialize progress bar properly for auto move
            self.update_custom_progress(0)
            self.update_idletasks()  # Force UI update

            # Launch the automation in a separate thread
            threading.Thread(
                target=self._run_stop_movement_worker,
                args=(mismatch_data,),
                daemon=True
            ).start()
            
        except Exception as e:
            print(f"âŒ Error starting auto move: {e}")
    
    def _check_shift_f7(self, event):
        """Check if shift-f7 is pressed to show/hide Auto Move button"""
        try:
            # Initialize flag if not exists
            if not hasattr(self, '_shift_f7_pressed'):
                self._shift_f7_pressed = False
            
            # Toggle Auto Move button visibility
            if hasattr(self, 'auto_move_button'):
                if self._shift_f7_pressed:
                    # Hide the Auto Move button
                    self.auto_move_button.pack_forget()
                    self._shift_f7_pressed = False
                    print("ðŸ” Auto Move button hidden")
                else:
                    # Show the Auto Move button (to the left of Copy Results button)
                    self.auto_move_button.pack(side="right", padx=(0, 5))
                    self._shift_f7_pressed = True
                    print("ðŸ” Auto Move button shown")
        except Exception as e:
            print(f"Error toggling Auto Move button: {e}")
        
        # Return "break" to prevent event propagation to other widgets
        return "break"
    
    def init_login_ui(self):
        if self.login_frame:
            self.clear_frame(self.login_frame)
        else:
            # 1) create a full-window transparent Canvas for the raindrop-ripple background
            import random, string
            self.login_canvas = tk.Canvas(self, bg="black", highlightthickness=0)  # Keep black for compatibility
            self.login_canvas.pack(expand=True, fill="both")

            # 2) layer your CTkFrame on top of that Canvas (centered) - transparent
            self.login_frame = ctk.CTkFrame(self.login_canvas, fg_color="transparent")
            self.login_frame.place(relx=0.5, rely=0.42, anchor="center")

            # 3) start the rain loop
            self._matrix_columns = int(self.winfo_screenwidth() / 20)
            self._matrix_drops   = [random.randint(-20, 0) for _ in range(self._matrix_columns)]
            # store the after() job so we can cancel it later
            self._matrix_job = self.after(50, self._draw_matrix)
        self.login_frame.configure(width=300)

        # â”€â”€ Carrier Logo & Title â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        import sys, os
        from PIL import Image
        from customtkinter import CTkImage

        # locate the bundled logo (works under PyInstaller)
        base_dir  = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
        logo_path = os.path.join(base_dir, "carrierx_logo.png")

        try:
            pil_img = Image.open(logo_path).convert("RGBA")
            
            # resize to max width 200px
            w, h    = pil_img.size
            new_w   = 200
            new_h   = int(h * (new_w / w))
            pil_img = pil_img.resize((new_w, new_h), Image.LANCZOS)
            
            # wrap in a CTkImage so CTkLabel won't warn
            ctk_logo = CTkImage(light_image=pil_img, size=(new_w, new_h))
            self.logo_img = ctk_logo
            logo_lbl = ctk.CTkLabel(self.login_frame, image=ctk_logo, text="")
        except Exception as e:
            # fallback to text if loading fails
            logo_lbl = ctk.CTkLabel(
                self.login_frame,
                text="Carrier",
                font=ctk.CTkFont(size=32, weight="bold"),
                text_color=BRAND_PURPLE
            )

        logo_lbl.pack(pady=(20, 10))

        # â”€â”€ Username & Password Entries â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.username_entry = ctk.CTkEntry(
            self.login_frame, placeholder_text="Carrier ID",
            font=ctk.CTkFont(size=14), width=200,
            fg_color="white", text_color="black", border_color=BRAND_ORANGE
        )
        self.username_entry.pack(pady=8)

        pw_frame = ctk.CTkFrame(self.login_frame, fg_color="black", width=200, height=30)
        pw_frame.pack(pady=8); pw_frame.pack_propagate(False)
        self.password_entry = ctk.CTkEntry(
            pw_frame, placeholder_text="Password", show="*",
            font=ctk.CTkFont(size=14), fg_color="white",
            text_color="black", border_color=BRAND_ORANGE
        )
        self.password_entry.pack(fill="both", expand=True)
        self.password_entry.bind("<Return>", lambda e: self.show_dispatch_ui())

        self.show_password_button = ctk.CTkButton(
            pw_frame, text="ðŸ‘", width=22, height=22,
            fg_color="transparent", hover_color=BRAND_ORANGE,
            command=self.toggle_password_visibility,
            font=ctk.CTkFont(size=14),
            border_width=0,
            corner_radius=4
        )
        self.show_password_button.place(relx=0.92, rely=0.5, anchor="center")

        # â”€â”€ Login Button â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.login_button = ctk.CTkButton(
            self.login_frame,
            text="Login",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self.show_dispatch_ui
        )
        self.login_button.pack(pady=(20, 10))

        # â”€â”€ Loading Animation with Carrier Truck (hidden until login starts) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Custom progress bar with Carrier truck and road
        self.login_progress_frame = ctk.CTkFrame(self.login_frame, fg_color="black", height=50)
        # Don't pack yet - will be packed when login starts
        
        # Road background (black with orange lines)
        self.login_road_canvas = tk.Canvas(
            self.login_progress_frame,
            bg="black",  # Black road to match UI
            height=30,
            highlightthickness=0
        )
        self.login_road_canvas.pack(pady=5, fill="x", padx=5)
        
        # Create progress overlay canvas
        self.login_progress_canvas = tk.Canvas(
            self.login_progress_frame,
            bg="black",
            height=30,
            highlightthickness=0
        )
        self.login_progress_canvas.place(x=5, y=5, relwidth=0, height=30)
        
        # Load Carrier truck logo for login
        try:
            from PIL import Image, ImageTk
            # Use bundled logo file (works under PyInstaller)
            base_dir = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
            truck_logo_path = os.path.join(base_dir, "TRUCK_LOGO.png")
            truck_img = Image.open(truck_logo_path)
            # Resize to fit login progress bar height
            truck_img = truck_img.resize((40, 25), Image.LANCZOS)
            self.login_truck_photo = ImageTk.PhotoImage(truck_img)
            
            # Create truck on login road canvas (initially hidden)
            self.login_truck_id = self.login_road_canvas.create_image(
                10, 15,  # Start position
                image=self.login_truck_photo,
                anchor="w",
                state="hidden"  # Initially hidden
            )
        except Exception as e:
            print(f"Could not load truck logo for login: {e}")
            # Fallback: create a simple truck shape
            self.login_truck_id = self.login_road_canvas.create_rectangle(
                10, 8, 50, 22,
                fill=BRAND_ORANGE,
                outline="white",
                state="hidden"  # Initially hidden
            )
        
        # Login truck animation variables
        self.login_truck_breathing_active = False
        self.login_truck_breathing_alpha = 1.0
        self.login_truck_breathing_direction = -1
        
        # Login progress text overlay
        self.login_progress_text = ctk.CTkLabel(
            self.login_progress_frame,
            text="",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=BRAND_ORANGE
        )
        self.login_progress_text.place(relx=0.5, rely=0.5, anchor="center")
        # â”€â”€ Status Text (hidden until login starts) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.login_status_label = ctk.CTkLabel(
            self.login_frame,
            text="",  # starts empty
            font=ctk.CTkFont(size=14),
            text_color="white"
        )
        # note: we DO NOT pack() this label here, so it stays hidden

        # Login breathing animation variables
        self.login_breathing_active = False
        self.login_breathing_alpha = 1.0
        self.login_breathing_direction = -1

    def _draw_matrix(self):
        import random, string
        # clear last frame
        self.login_canvas.delete("matrix")

        for i, drop in enumerate(self._matrix_drops):
            # pick a random char + Carrier palette color
            char  = random.choice(string.ascii_letters + string.digits)
            color = random.choice([BRAND_PURPLE, BRAND_ORANGE, "#FFFFFF"])
            x = i * 20
            y = drop * 20

            # draw
            self.login_canvas.create_text(
                x, y,
                text=char,
                fill=color,
                font=("Consolas", 16),
                tag="matrix"
            )

            # move drop
            self._matrix_drops[i] += 1
            # reset when off bottom
            if self._matrix_drops[i] * 20 > self.login_canvas.winfo_height():
                self._matrix_drops[i] = random.randint(-20, 0)

        # schedule next frame
        self._matrix_job = self.after(50, self._draw_matrix)

    def draw_login_road_lines(self, progress_width=0):
        """Draw dashed road lines on the login canvas - only show completed portion"""
        canvas_width = self.login_road_canvas.winfo_width()
        if canvas_width <= 1:  # Canvas not yet sized
            self.login_road_canvas.after(100, lambda: self.draw_login_road_lines(progress_width))
            return
        
        # Clear existing lines
        self.login_road_canvas.delete("login_road_line")
        
        # Only draw road lines up to the progress width
        y_center = 15
        dash_length = 15
        gap_length = 10
        
        x = 0
        while x < progress_width:
            # Calculate dash end position
            dash_end = min(x + dash_length, progress_width)
            
            # Only draw if dash is within progress area
            if x < progress_width:
                self.login_road_canvas.create_line(
                    x, y_center, dash_end, y_center,
                    fill=BRAND_ORANGE, width=2, tags="login_road_line"
                )
            x += dash_length + gap_length
    
    def update_login_progress(self, percentage):
        """Update the login progress bar with truck movement - thread-safe"""
        # Marshal to main thread if called from background thread
        if threading.current_thread() != threading.main_thread():
            self.after(0, lambda: self.update_login_progress(percentage))
            return
        
        # Show truck when process starts (at 1% or higher)
        if percentage >= 1.0 and not self.login_truck_breathing_active:
            self.login_road_canvas.itemconfig(self.login_truck_id, state="normal")
            self.login_truck_breathing_active = True
            # Only start breathing animation when actual progress begins (above 1%)
            if percentage > 1.0:
                self.start_login_truck_breathing()
        
        # Hide truck when process stops
        if percentage == 0 and self.login_truck_breathing_active:
            self.login_road_canvas.itemconfig(self.login_truck_id, state="hidden")
            self.login_truck_breathing_active = False
            # Reset percentage text color to default when breathing stops
            self.login_progress_text.configure(text_color=BRAND_ORANGE)
        
        # Update progress text (only show when process is running)
        if percentage > 0:
            self.login_progress_text.configure(text=f"{percentage:.1f}%")
        else:
            self.login_progress_text.configure(text="")
        
        # Calculate progress width
        canvas_width = self.login_road_canvas.winfo_width()
        if canvas_width <= 1:  # Canvas not yet sized
            return
        
        # Calculate progress width (progressive - only show completed portion)
        progress_width = (percentage / 100) * (canvas_width - 10)  # Leave 5px padding on each side
        
        # Update progress canvas width (this creates the progressive effect)
        self.login_progress_canvas.configure(width=progress_width)
        
        # Move truck to the end of the progress (spearhead position)
        # Only move truck if progress is actually advancing
        if progress_width > 0:
            truck_x = progress_width + 5  # Position truck just ahead of the progress
            if truck_x < 10:  # Keep truck at minimum position
                truck_x = 10
            
            # Move truck
            self.login_road_canvas.coords(self.login_truck_id, truck_x, 15)
        
        # Redraw road lines with current progress width
        self.draw_login_road_lines(progress_width)
    
    def start_login_truck_breathing(self):
        """Start the login truck breathing animation"""
        if not self.login_truck_breathing_active:
            return
        
        # Update breathing alpha (faster breathing for visibility)
        self.login_truck_breathing_alpha += self.login_truck_breathing_direction * 0.08
        
        # Reverse direction at boundaries
        if self.login_truck_breathing_alpha <= 0.2:
            self.login_truck_breathing_alpha = 0.2
            self.login_truck_breathing_direction = 1
        elif self.login_truck_breathing_alpha >= 1.0:
            self.login_truck_breathing_alpha = 1.0
            self.login_truck_breathing_direction = -1
        
        # Apply breathing effect to truck and percentage text
        try:
            # Get current truck position
            coords = self.login_road_canvas.coords(self.login_truck_id)
            if coords:
                # Handle both image (2 coords) and rectangle (4 coords) trucks
                if len(coords) == 2:  # Image truck
                    x, y = coords
                    # Add more visible breathing movement
                    breathing_offset = (self.login_truck_breathing_alpha - 0.5) * 4  # Doubled movement
                    self.login_road_canvas.coords(self.login_truck_id, x + breathing_offset, y)
                elif len(coords) == 4:  # Rectangle truck
                    x1, y1, x2, y2 = coords
                    width = x2 - x1
                    height = y2 - y1
                    # Add more visible breathing movement
                    breathing_offset = (self.login_truck_breathing_alpha - 0.5) * 4  # Doubled movement
                    self.login_road_canvas.coords(self.login_truck_id, x1 + breathing_offset, y1, x2 + breathing_offset, y2)
        except Exception as e:
            # Silently handle any coordinate errors
            pass
        
        # Apply breathing effect to percentage text color (more dramatic)
        carrierx_orange_rgb = (255, 102, 0)  # #FF6600
        new_color = tuple(int(c * self.login_truck_breathing_alpha) for c in carrierx_orange_rgb)
        breathing_color = f"#{new_color[0]:02x}{new_color[1]:02x}{new_color[2]:02x}"
        self.login_progress_text.configure(text_color=breathing_color)
        
        # Schedule next breathing frame (faster updates for more visible breathing)
        if self.login_truck_breathing_active:
            self.after(30, self.start_login_truck_breathing)  # Faster breathing updates

    def start_login_text_breathing(self):
        """Start the login status text breathing animation"""
        if not self.login_breathing_active:
            return
        
        # Update breathing alpha (faster breathing for visibility)
        self.login_breathing_alpha += self.login_breathing_direction * 0.1
        
        # Reverse direction at boundaries
        if self.login_breathing_alpha <= 0.3:
            self.login_breathing_alpha = 0.3
            self.login_breathing_direction = 1
        elif self.login_breathing_alpha >= 1.0:
            self.login_breathing_alpha = 1.0
            self.login_breathing_direction = -1
        
        # Apply breathing effect to status text color
        carrierx_orange_rgb = (255, 102, 0)  # #FF6600
        new_color = tuple(int(c * self.login_breathing_alpha) for c in carrierx_orange_rgb)
        breathing_color = f"#{new_color[0]:02x}{new_color[1]:02x}{new_color[2]:02x}"
        self.login_status_label.configure(text_color=breathing_color)
        
        # Schedule next breathing frame
        if self.login_breathing_active:
            self.after(50, self.start_login_text_breathing)

    def toggle_password_visibility(self):
        if self.password_entry.cget("show") == "*":
            self.password_entry.configure(show="")
            self.show_password_button.configure(text="ðŸ™ˆ")
        else:
            self.password_entry.configure(show="*")
            self.show_password_button.configure(text="ðŸ‘ï¸")

    def check_access(self, username: str) -> bool:
        # â”€â”€â”€ Bypass GUI access check for special IDs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if username.strip() in BYPASS_ACCESS_IDS:
            return True
        
        logging.info(f"ðŸ” Starting auth for Carrier ID: {username}")
        driver_path = ChromeDriverManager().install()
        opts = Options()
        opts.add_argument("--headless")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--disable-gpu")
        service = Service(driver_path, log_path=os.devnull)
        driver  = webdriver.Chrome(service=service, options=opts)
        try:
            logging.info("âž¡ï¸  Loading EmployeeDirectory page")
            driver.get("https://internal.example.com")
            wait = WebDriverWait(driver, 10)

            # 1) search by Carrier ID
            logging.info("ðŸ” Locating search box")
            search = wait.until(EC.element_to_be_clickable((By.NAME, "Search")))
            search.clear()
            search.send_keys(username, Keys.RETURN)
            logging.info(f"âœï¸  Submitted search for {username}")

            # Try to capture the user's display name from the search results link
            try:
                name_link = wait.until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        f"//a[contains(@onclick, \"employeeId.value='{username}'\")]"
                    ))
                )
                raw_name = name_link.text.strip().replace('\u00A0', ' ')
                self._set_agent_display_name_from_full(raw_name)
            except TimeoutException:
                logging.warning("âš ï¸ Could not capture dispatcher name from EmployeeDirectory search results")

            # â€”â€”â€” Allow-list: grant access if searched user is one of these â€”â€”â€”
            try:
                # adjust this XPath to point at wherever EmployeeDirectory shows the user's name
                name_el = wait.until(EC.presence_of_element_located((
                    By.XPATH,
                    "//b[normalize-space(text())='Name:']/following-sibling::a"
                )))
                allowed_names = {
                    "Pamela M. Scarborough",
                    "Rickie Murata",
                    "Kurt T. Mahoney",
                    "Thomas Sonsini",
                }
                if name_el.text.strip() in allowed_names:
                    logging.info(f"âœ… Searched user is {name_el.text.strip()} â€” access granted")
                    return True
            except TimeoutException:
                # fall back to manager-chain logic if name field didn't appear
                pass

            # 2) define how to locate the Manager link
            logging.info("ðŸ” Preparing manager-link locator")
            manager_xpath = (
                "//b[normalize-space(text())='Manager:']"
                "/following-sibling::a"
            )

            # 3) follow up the chain up to 2 levels until we find an allowed manager
            logging.info("ðŸ” Checking manager chain (up to 2 levels)")
            for level in range(1, 3):
                logging.info(f"  â†ªï¸  Level {level}: locating manager link")
                mgr_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, manager_xpath))
                )
                mgr_name_raw = mgr_link.text.strip()
                # normalize NBSP â†’ regular space
                mgr_name = mgr_name_raw.replace('\u00A0', ' ')
                logging.info(f"  ðŸ‘¤ Found manager raw: {mgr_name_raw!r}, normalized: {mgr_name!r}")
                if mgr_name in {
                    "Pamela M. Scarborough",
                    "Rickie Murata",
                    "Kurt T. Mahoney",
                    "Thomas Sonsini",
                }:
                    logging.info(f"âœ… {mgr_name} verified â€” access granted")
                    return True
                logging.info(f"âŒ Not Pamela, drilling into {mgr_name}'s page")
                mgr_link.click()
                WebDriverWait(driver, 5).until(EC.staleness_of(mgr_link))

            # never encountered an allowed manager in the chain
            logging.error("â›” Allowed manager not found in manager chain â€” access denied")
            return False
        except TimeoutException:
            logging.exception("âš ï¸ Timeout during auth check")
            return False
        finally:
            logging.info("ðŸ›‘ Auth check complete, closing browser")
            driver.quit()

    # â”€â”€â”€ SSO / Identity Portal login helper â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def sso_navigate(self, driver, wait,
                      url="https://internal.example.com",
                      username=None, password=None):
        """
        Navigate to `url` and, if you land on the SSO/Identity Portal page,
        submit creds using the new two-step authentication flow and wait for the FRO menu to appear.
        """
        driver.get(url)

        # â€” Poll frequently for auto-login (automatic login via SSO/cookies can take ~10 seconds) â€”
        logging.info("ðŸ” Checking for auto-login (polling every 1 second for up to 15 seconds)...")
        for attempt in range(15):  # Check for up to 15 seconds
            if CarrierSeleniumMonitor.check_if_logged_in(driver):
                logging.info(f"âœ… Auto-login detected after {attempt + 1} seconds (URL contains /home and FRO menu found)")
                return  # Already logged in, skip login process
            time.sleep(1)
        
        # Final check after polling period
        if CarrierSeleniumMonitor.check_if_logged_in(driver):
            logging.info("âœ… Auto-login detected after polling period")
            return  # Already logged in, skip login process

        try:
            # Step 1: Enter username and click Next
            logging.info("ðŸ”„ Step 1: Entering username...")
            # Wait for username input field to appear using dynamic detection (with frequent auto-login checks)
            logging.info("ðŸ” Waiting for username input field to appear...")
            user = None
            start_time = time.time()
            timeout = 30
            check_interval = 1  # Check every 1 second
            
            while time.time() - start_time < timeout:
                # Check for auto-login first
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected during username field wait")
                    return  # Already logged in, skip login process
                
                # Try to find username field
                try:
                    user = CarrierSeleniumMonitor.find_username_input(driver, None)
                    if user and user.is_displayed():
                        break
                except:
                    pass
                
                time.sleep(check_interval)
            
            if not user:
                # Before raising error, final check for auto-login
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected (final check)")
                    return  # Already logged in, skip login process
                
                # Fallback: try original method
                try:
                    user = wait.until(CarrierSeleniumMonitor.UsernameFieldFound())
                except TimeoutException:
                    # Final check for auto-login before raising exception
                    if CarrierSeleniumMonitor.check_if_logged_in(driver):
                        logging.info("âœ… Auto-login detected during username wait timeout")
                        return  # Already logged in, skip login process
                    
                    # Fallback to waiting for any clickable input (original behavior)
                    logging.warning("âš ï¸ Could not dynamically find username input, trying fallback...")
                    try:
                        # Try common input ID patterns
                        for input_id in ["input44", "input28", "input"]:
                            try:
                                user = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
                                break
                            except:
                                continue
                        # If still not found, get first visible text input
                        if not user:
                            user = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
                    except:
                        # Final check for auto-login before raising exception
                        if CarrierSeleniumMonitor.check_if_logged_in(driver):
                            logging.info("âœ… Auto-login detected (fallback check)")
                            return  # Already logged in, skip login process
                        raise Exception("Could not find username input field")
            
            logging.info(f"âœ… Found username input: {user.get_attribute('id') or user.get_attribute('name') or 'unknown'}")
            user.clear()
            user.send_keys(username)
            
            # Look for Next button - try multiple possible selectors
            next_button = None
            next_selectors = [
                "input[type='submit']",
                "button[type='submit']", 
                "input[value*='Next']",
                "button:contains('Next')",
                ".next-button",
                "#next-button",
                "input[value='Next']",
                "button[value='Next']"
            ]
            
            for selector in next_selectors:
                try:
                    if selector.startswith("button:contains"):
                        # Use XPath for text content
                        next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]")))
                    else:
                        next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    break
                except TimeoutException:
                    continue
            
            if next_button:
                logging.info("ðŸ”„ Found Next button, clicking...")
                next_button.click()
                # Wait a moment for the page to transition
                time.sleep(2)
            else:
                logging.info("ðŸ”„ No Next button found, trying Enter key...")
                user.send_keys(Keys.RETURN)
                time.sleep(2)
            
            # Step 2: Enter password
            logging.info("ðŸ”„ Step 2: Entering password...")
            # Wait for password input field to appear using dynamic detection (with frequent auto-login checks)
            logging.info("ðŸ” Waiting for password input field to appear...")
            pw = None
            start_time = time.time()
            timeout = 30
            check_interval = 1  # Check every 1 second
            
            while time.time() - start_time < timeout:
                # Check for auto-login first
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected during password field wait")
                    return  # Already logged in, skip login process
                
                # Try to find password field
                try:
                    pw = CarrierSeleniumMonitor.find_password_input(driver, None)
                    if pw and pw.is_displayed():
                        break
                except:
                    pass
                
                time.sleep(check_interval)
            
            if not pw:
                # Before raising error, final check for auto-login
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected (final check)")
                    return  # Already logged in, skip login process
                
                # Fallback: try original method
                try:
                    pw = wait.until(CarrierSeleniumMonitor.PasswordFieldFound())
                except TimeoutException:
                    # Final check for auto-login before raising exception
                    if CarrierSeleniumMonitor.check_if_logged_in(driver):
                        logging.info("âœ… Auto-login detected during password wait timeout")
                        return  # Already logged in, skip login process
                    
                    # Fallback to waiting for password input (original behavior)
                    logging.warning("âš ï¸ Could not dynamically find password input, trying fallback...")
                    try:
                        # Try common input ID patterns
                        for input_id in ["input70", "input54", "input"]:
                            try:
                                pw = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
                                break
                            except:
                                continue
                        # If still not found, get password type input
                        if not pw:
                            pw = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
                    except:
                        # Final check for auto-login before raising exception
                        if CarrierSeleniumMonitor.check_if_logged_in(driver):
                            logging.info("âœ… Auto-login detected (fallback check)")
                            return  # Already logged in, skip login process
                        raise Exception("Could not find password input field")
            
            logging.info(f"âœ… Found password input: {pw.get_attribute('id') or pw.get_attribute('name') or 'unknown'}")
            pw.clear()
            pw.send_keys(password + Keys.RETURN)
            
            # Confirm login by waiting for the DSM header instead
            CarrierSeleniumMonitor.ensure_home_loaded(driver)
        except TimeoutException:
            # either no login prompt or already logged in
            pass
    
    def _async_access_check(self):
        """
        Runs in a background thread: checks only the manager-chain access,
        then marshals UI updates back onto the Tk mainloop via self.after().
        """
        # Define update_progress function for all users
        def update_progress(progress):
            self.after(0, lambda: self.update_login_progress(progress))
        
        import time
        
        # Check if user is in bypass list first
        if self.username.strip() in BYPASS_ACCESS_IDS:
            # Bypass user - show quick progress and grant access immediately
            
            # Quick progress for bypass users (1-100% in 1 second)
            for i in range(1, 101):
                update_progress(i)
                time.sleep(0.01)  # Very fast progress
            
            allowed = True
        else:
            # Regular user - START AUTH IMMEDIATELY and show concurrent progress
            allowed_holder = {"done": False, "value": False}
            
            def do_check():
                try:
                    allowed_holder["value"] = self.check_access(self.username)
                finally:
                    allowed_holder["done"] = True
            
            import threading
            t = threading.Thread(target=do_check, daemon=True)
            t.start()
            
            # Animate progress while auth runs; pace to ~12 seconds to reach 95%
            progress = 1
            update_progress(progress)
            start_ts = time.time()
            target_secs = 27.0
            while not allowed_holder["done"]:
                elapsed = time.time() - start_ts
                # proportionally map elapsed time to 1..95
                pct = 1 + int(min(94, max(0, (elapsed / target_secs) * 94)))
                if pct > progress:
                    progress = pct
                    update_progress(progress)
                time.sleep(0.05)
            
            allowed = allowed_holder["value"]

        # Final completion (91-100%) - Much faster
        start = 91
        # If progress already reached 95, finish from there
        try:
            current = max(start, int(progress) if 'progress' in locals() else start)
        except Exception:
            current = start
        for i in range(current, 101):
            update_progress(i)
            time.sleep(0.01)  # 5x faster final completion
        def on_done():
            # Immediately stop all animations and hide UI elements
            self.login_breathing_active = False
            self.update_login_progress(0)  # Reset to 0 to hide truck
            self.login_progress_frame.pack_forget()
            self.login_status_label.pack_forget()

            # Re enable the inputs
            self.login_button.configure(state="normal")
            self.username_entry.configure(state="normal")
            self.password_entry.configure(state="normal")

            if not allowed:
                # Manager-chain check failed â†’ send user back to login
                messagebox.showerror(
                    "Authentication Failed",
                    "You do not have access!.\nIf you believe this is an error please contact your manager."
                )
                return

            # 1) stop the matrix animation loop
            if hasattr(self, '_matrix_job'):
                self.after_cancel(self._matrix_job)
            # 2) destroy the background canvas entirely
            self.login_canvas.destroy()

            # 3) tear down login frame
            self.login_frame.destroy()
            self.login_frame = None

            # 4) show main UI
            self.init_dispatch_ui()

        # Schedule on_done() on the main thread
        self.after(0, on_done)

    def check_sso(self, username: str, password: str) -> bool:
        """
        Returns True if SSO login succeeds against the Identity Portal agentless-DS-SSO page;
        returns False (staying on login) if it fails.
        """
        LOGIN_URL = "https://internal.example.com"
        # mirror your Auto-Clear Chrome options
        driver_path = ChromeDriverManager().install()
        service     = Service(driver_path, log_path=os.devnull)
        opts        = Options()
        opts.add_argument("--disable-extensions")
        opts.add_argument("--headless")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--disable-background-networking")
        opts.add_argument("--disable-gcm")
        opts.add_argument("--log-level=3")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        # Enable Chrome performance logging for Network events (to capture CSV response)
        try:
            opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
        except Exception:
            pass
        driver  = webdriver.Chrome(service=service, options=opts)

        try:
            driver.get(LOGIN_URL)
            # Wait longer in headed mode for testing - can see what's happening
            short_wait = WebDriverWait(driver, 30)
            try:
                # Dynamically check if login form exists
                # Try to find username input - if found, login form exists
                username_input = None
                for attempt in range(5):  # Try up to 5 times quickly
                    try:
                        username_input = CarrierSeleniumMonitor.find_username_input(driver, short_wait)
                        if username_input and username_input.is_displayed():
                            break
                    except:
                        pass
                    time.sleep(0.2)
                
                if username_input and username_input.is_displayed():
                    # form is there â†’ do full MFA-aware login
                    helper = CarrierSeleniumMonitor(
                        username=username, password=password,
                        dispatch_location="", stop_event=threading.Event(),
                        message_types=[],
                        mfa_factor_callback=self.prompt_mfa_factor,
                        mfa_code_callback  =self.prompt_mfa_code
                    )
                    helper.sso_navigate(driver, short_wait, LOGIN_URL, username, password)
                else:
                    # no form â†’ assume already SSO'd in
                    pass
            except (TimeoutException, Exception):
                # no form â†’ assume already SSO'd in
                pass

            # confirm we've landed in the app
            CarrierSeleniumMonitor.ensure_home_loaded(driver)
            # Add delay in headed mode so you can see the result before browser closes
            logging.info("âœ… Login successful - keeping browser open for 10 seconds for observation...")
            time.sleep(10)
            return True

        except TimeoutException:
            # Add delay in headed mode so you can see what happened
            logging.warning("âš ï¸ Timeout during login - keeping browser open for 10 seconds for observation...")
            time.sleep(10)
            return False

        finally:
            driver.quit()    

    def show_dispatch_ui(self):
        # collect credentials
        self.username = self.username_entry.get().strip()
        self.password = self.password_entry.get().strip()
        if not self.username or not self.password:
            messagebox.showerror("Missing Info", "Enter both username and password.")
            return

        # â”€â”€ Disable inputs so user can't click twice â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.login_button.configure(state="disabled")
        self.username_entry.configure(state="disabled")
        self.password_entry.configure(state="disabled")

        # â€¦ after disabling inputs â€¦
        self.login_status_label.configure(text="Authenticatingâ€¦")
        self.login_status_label.pack(pady=(0, 5))

        # Start the text breathing animation
        self.login_breathing_active = True
        self.start_login_text_breathing()

        # â”€â”€ Show & start truck progress bar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.login_progress_frame.pack(pady=(5, 0), fill="x", padx=20)
        self.update_login_progress(1)  # Start at 1% to show truck

        # â”€â”€ Spin off the access check so the UI stays responsive â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        threading.Thread(target=self._async_access_check, daemon=True).start()

    def init_dispatch_ui(self):
        if self.dispatch_frame:
            self.clear_frame(self.dispatch_frame)
        else:
            self.dispatch_frame = ctk.CTkFrame(self, fg_color="black")
            self.dispatch_frame.pack(expand=True, fill="both")

        # Create a container frame for tabs to allow second row positioning
        self.tabs_container = ctk.CTkFrame(self.dispatch_frame, fg_color="transparent")
        self.tabs_container.pack(pady=10)
        
        self.tab = ctk.CTkTabview(self.tabs_container, width=580, height=620, fg_color="black")
        self.tab.pack()
        try:
            seg = self.tab._segmented_button
            seg.configure(fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE, selected_color=BRAND_PURPLE)
            for ch in seg.winfo_children():
                ch.configure(hover_color=BRAND_PURPLE, selected_color=BRAND_PURPLE)
        except:
            pass

        self.tab.add("Auto-Read")
        self.tab.add("Auto-Clear")
        self.tab.add("SoF+")

        # Initialize widget system
        self.widget_panel_visible = False
        self.create_widget_button()

        # Auto-Sort tab is now enabled
        # try:
        #     seg = tab._segmented_button
        #     for btn in seg.winfo_children():
        #         if btn.cget("text") == "Auto-Sort":
        #             btn.configure(state="disabled")
        # except Exception:
        #     pass

        # â”€â”€â”€ Utility Tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.tab.add("Utility")
        # â”€â”€â”€ Action Log Tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.tab.add("Action Log")
        # â”€â”€â”€ Reports Tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.tab.add("Reports")
        # â”€â”€â”€ Settings Tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.tab.add("Settings")
        
        # Store previous tab for "coming soon" functionality
        self._previous_tab = "Auto-Read"
        
        # Bind tab change event to handle widget visibility
        # Wrap in lambda to get current tab, as CTkTabView doesn't pass it as argument
        self.tab.configure(command=lambda: self._handle_tab_change(self.tab.get()))
        
        # Defer heavy tab content building to make tabs appear instantly
        self._tabs_initialized = getattr(self, "_tabs_initialized", False)
        if not self._tabs_initialized:
            self._tabs_initialized = True
            # Build all tab content asynchronously after UI is shown
            self.after(0, self._build_all_tabs_content)
        
        # Update widget button visibility after UI is set up
        self.after(100, self.update_widget_button_visibility)
    
    def _handle_tab_change(self, selected_tab):
        """Handle tab changes, including widget visibility"""
        try:
            # Update widget button visibility
            self.update_widget_button_visibility()
            
            # Store the current tab as previous
            self._previous_tab = selected_tab
        except Exception as e:
            # Log errors during tab change but don't spam
            if not hasattr(self, '_last_tab_change_error') or self._last_tab_change_error != str(e):
                logging.debug(f"Error during tab change: {e}")
                self._last_tab_change_error = str(e)
    
    def _build_all_tabs_content(self):
        """Build all tab content asynchronously to keep tabs appearing instantly"""
        # â”€â”€â”€ Utility Tab Content â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        util_tab = self.tab.tab("Utility")

        # Dropbox Code Generator UI
        util_frame = ctk.CTkFrame(
            util_tab,
            fg_color="black",
            border_width=2,
            border_color=BRAND_ORANGE,
            corner_radius=0
        )
        util_frame.pack(padx=25, pady=(15, 25), fill="x")

        # Title
        ctk.CTkLabel(
            util_frame,
            text="Generate Dropbox Code",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=BRAND_ORANGE
        ).grid(row=0, column=0, columnspan=6, pady=(8,10))

        # Row 1: Dropbox ID, Driver Type dropdown
        util_frame.grid_columnconfigure(0, minsize=95)
        ctk.CTkLabel(util_frame, text="Dropbox ID:", text_color=BRAND_ORANGE)\
            .grid(row=1, column=0, sticky="e", padx=(10,8))
        self.dropbox_id_entry = ctk.CTkEntry(util_frame, width=100)
        self.dropbox_id_entry.grid(row=1, column=1, sticky="w")

        util_frame.grid_columnconfigure(2, minsize=90)
        ctk.CTkLabel(util_frame, text="Driver Type:", text_color=BRAND_ORANGE)\
            .grid(row=1, column=2, sticky="e", padx=(15,8))
        self.driver_type_var = tk.StringVar(value="Courier")
        self.driver_type_menu = ctk.CTkOptionMenu(
            util_frame,
            variable=self.driver_type_var,
            values=["Courier", "Service Provider Driver"],
            width=175,
            fg_color=BRAND_ORANGE,
            button_color=BRAND_ORANGE,
            button_hover_color=BRAND_PURPLE,
            text_color="white",
            dropdown_fg_color="black",
            dropdown_hover_color=BRAND_PURPLE,
            command=self._on_driver_type_change
        )
        self.driver_type_menu.grid(row=1, column=3, sticky="w", padx=(0, 15))

        # Row 2: Courier fields (Employee ID) + Generate - shown when Courier selected
        self.employee_id_label = ctk.CTkLabel(util_frame, text="Employee ID:", text_color=BRAND_ORANGE)
        self.employee_id_label.grid(row=2, column=0, sticky="e", padx=(10,5), pady=(6,6))
        self.employee_id_entry = ctk.CTkEntry(util_frame, width=100)
        self.employee_id_entry.grid(row=2, column=1, sticky="w", pady=(6,6))

        # Row 2: Service Provider Driver fields + Generate - shown when Service Provider Driver selected
        self.driver_vendor_label = ctk.CTkLabel(util_frame, text="Vendor ID:", text_color=BRAND_ORANGE)
        self.driver_vendor_label.grid(row=2, column=0, sticky="e", padx=(10,5), pady=(6,6))
        self.driver_vendor_id_entry = ctk.CTkEntry(util_frame, width=100)
        self.driver_vendor_id_entry.grid(row=2, column=1, sticky="w", pady=(6,6))
        self.csa_label = ctk.CTkLabel(util_frame, text="CSA #:", text_color=BRAND_ORANGE)
        self.csa_label.grid(row=2, column=2, sticky="e", padx=(15,5), pady=(6,6))
        self.csa_number_entry = ctk.CTkEntry(util_frame, width=100)
        self.csa_number_entry.grid(row=2, column=3, sticky="w", pady=(6,6))

        self.driver_vendor_label.grid_remove()
        self.driver_vendor_id_entry.grid_remove()
        self.csa_label.grid_remove()
        self.csa_number_entry.grid_remove()

        # Reason dropdown
        ctk.CTkLabel(util_frame, text="Reason:", text_color=BRAND_ORANGE)\
            .grid(row=3, column=0, sticky="e", padx=(10,5), pady=(8,5))
        self.reason_var = tk.StringVar(value="Select One:")
        self.reason_menu = ctk.CTkOptionMenu(
            util_frame,
            variable=self.reason_var,
            values=[
                "Handheld won't open dropbox",
                "LEO control box battery issue",
                "LEO lock battery doesn't work",
                "No PowerPad",
                "Drop Box Battery is dead",
                "Used Handheld Device and Door did not open",
                "Other"
            ],
            width=240,
            fg_color=BRAND_ORANGE,
            button_color=BRAND_ORANGE,
            button_hover_color=BRAND_PURPLE,
            text_color="white",
            dropdown_fg_color="black",
            dropdown_hover_color=BRAND_PURPLE,
            command=self._on_reason_change
        )
        self.reason_menu.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(8,5))

        # Other-Reason (hidden until "Other" is selected)
        self.other_reason_label = ctk.CTkLabel(util_frame, text="Other Reason:", text_color=BRAND_ORANGE)
        self.other_reason_entry = ctk.CTkEntry(util_frame, width=300)
        self.other_reason_label.grid(row=4, column=0, sticky="e", padx=(10,5), pady=(5,5))
        self.other_reason_entry.grid(row=4, column=1, columnspan=4, sticky="w", pady=(5,5))
        self.other_reason_label.grid_remove()
        self.other_reason_entry.grid_remove()

        # Output (row 4 when Other hidden, row 5 when Other shown - set by _on_reason_change)
        self.output_label = ctk.CTkLabel(util_frame, text="Output:", text_color=BRAND_ORANGE)
        self.output_label.grid(row=4, column=0, sticky="e", padx=(10,5), pady=(8,8))

        self.output_entry = ctk.CTkEntry(util_frame, width=240, state="readonly")
        self.output_entry.grid(row=4, column=1, columnspan=2, sticky="ew", pady=(8,8))

        self.copy_button = ctk.CTkButton(
            util_frame,
            text="Copy",
            width=70, height=28,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=self._copy_dropbox_code
        )
        self.copy_button.grid(row=4, column=3, padx=(8,10), pady=(8,8), sticky="w")

        util_frame.grid_columnconfigure(1, weight=1)
        util_frame.grid_columnconfigure(4, weight=1)

        self.generate_button = ctk.CTkButton(
            util_frame,
            text="Generate",
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self._generate_dropbox_code_threaded
        )
        self.generate_button.grid(row=5, column=1, columnspan=2, pady=(12, 18), sticky="", padx=(25, 0))

        # â”€â”€â”€ RoutR Chrome Extension (Utility) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        ext_frame = ctk.CTkFrame(
            util_tab,
            fg_color="black",
            border_width=2,
            border_color=BRAND_ORANGE,
            corner_radius=0,
        )
        # Center the RoutR Chrome Extension box within the Utility tab
        ext_frame.pack(padx=25, pady=(0, 25), anchor="center")
        # Center contents inside the frame
        ext_frame.grid_columnconfigure(0, weight=1)
        ext_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            ext_frame,
            text="RoutR Chrome Extension",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=BRAND_ORANGE,
        ).grid(row=0, column=0, columnspan=2, pady=(8, 10), sticky="n")
        ctk.CTkLabel(
            ext_frame,
            text="Adds your CXPC signature to DSM messages in your browser when ROUTR is open.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).grid(row=1, column=0, columnspan=2, padx=(10, 10), pady=(0, 8))
        ctk.CTkButton(
            ext_frame,
            text="Install RoutR Chrome Extension",
            width=220,
            height=28,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self._on_install_routr_extension,
        ).grid(row=2, column=0, padx=(10, 10), pady=(0, 10))

        # Slider-style switch for enabling/disabling the extension (default ON)
        self.routr_extension_var = tk.BooleanVar(value=True)
        self.routr_extension_switch = ctk.CTkSwitch(
            ext_frame,
            text="Extension",
            variable=self.routr_extension_var,
            onvalue=True,
            offvalue=False,
            command=self._on_toggle_routr_extension,
            progress_color="#00FF00",  # green when ON
            fg_color="#FF0000",        # red when OFF
        )
        self.routr_extension_switch.grid(row=2, column=1, padx=(10, 10), pady=(0, 10))

        # â”€â”€â”€ Action Log Tab Content â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log_tab = self.tab.tab("Action Log")
        
        # â”€â”€â”€ Auto-Sort Tab (will be created when Shift+F6 is pressed) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Don't add the tab initially - it will be added dynamically when Shift+F6 is pressed
        # Text widget for action logs
        # â”€â”€â”€ make a scrollable Action Log â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log_container = ctk.CTkFrame(log_tab, fg_color="black")
        log_container.pack(fill="both", expand=True, padx=10, pady=10)

        # â”€â”€â”€ Text widget â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.log_text = ctk.CTkTextbox(
            log_container, width=560, height=550,
            fg_color="black", text_color="white"
        )
        self.log_text.configure(state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)

        # â”€â”€â”€ Vertical scrollbar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.log_scrollbar = ctk.CTkScrollbar(
            log_container, orientation="vertical",
            command=self.log_text.yview
        )
        self.log_scrollbar.pack(side="right", fill="y")

        # â”€â”€â”€ Hook scrollbar to text â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.log_text.configure(yscrollcommand=self.log_scrollbar.set)

        # --- Auto-Read Tab Content ---
        read = self.tab.tab("Auto-Read")
        ctk.CTkLabel(read, text="Enter Dispatch Location",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=BRAND_ORANGE).pack(pady=(10, 15))
        self.dispatch_entry = ctk.CTkEntry(
            read, font=ctk.CTkFont(size=14), width=60,
            fg_color="white", border_color=BRAND_ORANGE, text_color="black"
        )
        self.dispatch_entry.pack(pady=8)
        self.dispatch_entry.bind("<Return>", lambda e: self.add_location(False))
        self.dispatch_entry.bind("<KeyRelease>", lambda e: self.convert_to_uppercase())
        cbf = ctk.CTkFrame(read, fg_color="black"); cbf.pack(pady=(5, 15))
        self.pickup_reminder_var = ctk.BooleanVar(value=True)
        self.early_pu_var        = ctk.BooleanVar(value=True)
        self.no_pickup_list_var  = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(cbf, text="Pickup Reminder", variable=self.pickup_reminder_var,
                        fg_color=BRAND_ORANGE, text_color="white", hover_color=BRAND_PURPLE)\
            .grid(row=0, column=0, padx=10)
        self.reply_on_pr_var = ctk.BooleanVar(value=False)
        switch = ctk.CTkSwitch(
            cbf,
            text="Pickup Reminder Reply",
            variable=self.reply_on_pr_var,
            onvalue=True, offvalue=False,
            progress_color="#00FF00",  # green when on
            fg_color="#FF0000"         # red when off
        )
        # Move it under the Pickup Reminder checkbox (column=0) and align left
        switch.grid(row=1, column=0, sticky="w", pady=(5,0))
        ToolTip(switch, "When ON/Green, Pickup Reminders will be sent back to the courier with the stop info.")    
        # â”€â”€â”€ Early PU checkbox & reply toggle â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        ctk.CTkCheckBox(
            cbf, text="Early PU", variable=self.early_pu_var,
            fg_color=BRAND_ORANGE, text_color="white", hover_color=BRAND_PURPLE
        ).grid(row=0, column=1, padx=10)

        self.reply_on_epu_var = ctk.BooleanVar(value=False)
        switch_epu = ctk.CTkSwitch(
            cbf,
            text="Early PU Reply",
            variable=self.reply_on_epu_var,
            onvalue=True, offvalue=False,
            progress_color="#00FF00", fg_color="#FF0000"
        )
        switch_epu.grid(row=1, column=1, sticky="w", pady=(5,0))
        ToolTip(switch_epu, "When ON/Green, Early PU messages will get a follow-up question.")
        ctk.CTkCheckBox(cbf, text="No Pickup List", variable=self.no_pickup_list_var,
                        fg_color=BRAND_ORANGE, text_color="white", hover_color=BRAND_PURPLE)\
            .grid(row=0, column=2, padx=10)
        # â”€â”€â”€ Auto-WalkUp toggle â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.auto_walkup_var = ctk.BooleanVar(value=False)
        switch_walkup = ctk.CTkSwitch(
            cbf,
            text="Auto-WalkUp",
            variable=self.auto_walkup_var,
            onvalue=True, offvalue=False,
            progress_color="#00FF00",  # green when ON
            fg_color="#FF0000"         # red when OFF
        )
        switch_walkup.grid(row=1, column=2, sticky="w", pady=(5, 0))
        ToolTip(switch_walkup, "When ON/Green, Walk Ups will be sent to the couriers when they request it via #WU or #walkup.")  
        btn = ctk.CTkButton(read, text="Add Location",
                            font=ctk.CTkFont(size=14, weight="bold"),
                            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                            command=lambda: self.add_location(False))
        btn.pack(pady=15)
        if not self.locations_container:
            self.locations_container = ctk.CTkFrame(read, fg_color="black")
            self.locations_container.pack(pady=(10, 0), fill="x")

        # --- Auto-Clear Tab ---
        clear_tab = self.tab.tab("Auto-Clear")
        ctk.CTkLabel(clear_tab, text="Enter Dispatch Location",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=BRAND_ORANGE).pack(pady=(10, 15))
        dr = ctk.CTkFrame(clear_tab, fg_color="black"); dr.pack(pady=(0, 15))

        self.clear_dispatch_entry = ctk.CTkEntry(
            dr, font=ctk.CTkFont(size=14), width=60,
            fg_color="white", border_color=BRAND_ORANGE, text_color="black"
        )
        self.clear_dispatch_entry.pack(side="left", padx=(0, 10))
        self.clear_dispatch_entry.bind("<Return>", lambda e: self.add_clear_location(False))
        self.clear_dispatch_entry.bind("<KeyRelease>", lambda e: self.convert_to_uppercase_clear())

        rf = ctk.CTkFrame(clear_tab, fg_color="black"); rf.pack(pady=(5, 15))
        ctk.CTkLabel(rf, text="Clear Reply",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=BRAND_ORANGE).pack(side="left", padx=10)

        # Dropdown of common clear messages + Custom option
        self.clear_reply_entry = ctk.CTkEntry(
            rf, font=ctk.CTkFont(size=14), width=260,
            fg_color="white", border_color=BRAND_ORANGE, text_color="black"
        )
        default_clear = "Clear! Have a great day!"
        self.clear_reply_entry.insert(0, default_clear)
        # Do NOT pack yet; we only show this box when 'Custom' is selected.
        # allow Enter in the reply box to add the clear-location (when visible)
        self.clear_reply_entry.bind("<Return>", lambda e: self.add_clear_location(False))
        self.clear_reply_choice_var = tk.StringVar(value=default_clear)
        self.clear_reply_menu = ctk.CTkOptionMenu(
            rf,
            variable=self.clear_reply_choice_var,
            values=[
                "Clear! Have a great day!",
                "Clear! Have a great night!",
                "Clear! Have a great weekend!",
                "Custom",
            ],
            width=210,
            fg_color=BRAND_ORANGE,
            button_color=BRAND_ORANGE,
            button_hover_color=BRAND_PURPLE,
            text_color="white",
            dropdown_fg_color="black",
            dropdown_hover_color=BRAND_PURPLE,
            command=self._on_clear_reply_choice_change,
        )
        self.clear_reply_menu.pack(side="left", padx=(10, 10))

        # Initialize UI state so that the custom box is hidden
        self._on_clear_reply_choice_change(default_clear)

        # Row: "Sending clear message as ..." + Change button
        sig_frame = ctk.CTkFrame(clear_tab, fg_color="black")
        sig_frame.pack(pady=(0, 4), fill="x")

        sig_text = (
            f"Sending clear message as -CXPC {self.agent_display_name}"
            if self.agent_display_name
            else "Sending clear message without CXPC name (login or save a CXPC Name)."
        )
        self.clear_signature_label = ctk.CTkLabel(
            sig_frame,
            text=sig_text,
            font=ctk.CTkFont(size=12),
            text_color="white",
            justify="center",
        )
        self.clear_signature_label.pack(side="top", padx=(10, 6), pady=(0, 2))

        change_btn = ctk.CTkButton(
            sig_frame,
            text="Change",
            width=70,
            height=24,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self._on_toggle_clear_name_editor,
        )
        change_btn.pack(side="top", pady=(0, 4))

        # Manual override frame for dispatcher name (initially hidden)
        self.clear_name_frame = ctk.CTkFrame(clear_tab, fg_color="black")

        ctk.CTkLabel(
            self.clear_name_frame,
            text="CXPC Name:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=BRAND_ORANGE,
        ).pack(side="left", padx=(10, 6))

        self.clear_name_entry = ctk.CTkEntry(
            self.clear_name_frame,
            font=ctk.CTkFont(size=12),
            width=180,
            fg_color="white",
            border_color=BRAND_ORANGE,
            text_color="black",
        )
        # Pre-fill with current name if we already have one
        if self.agent_display_name:
            self.clear_name_entry.insert(0, self.agent_display_name)
        self.clear_name_entry.pack(side="left", padx=(0, 6))

        name_save_btn = ctk.CTkButton(
            self.clear_name_frame,
            text="Save Name",
            width=90,
            height=26,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self._on_save_clear_name,
        )
        name_save_btn.pack(side="left", padx=(0, 10))

        self.clear_add_location_button = ctk.CTkButton(
            clear_tab,
            text="Add Location",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=lambda: self.add_clear_location(False),
        )
        self.clear_add_location_button.pack(pady=15)

        if not self.clear_locations_container:
            # Scrollable container for per-location Auto-Clear monitors
            self.clear_locations_container = ctk.CTkScrollableFrame(
                clear_tab,
                fg_color="black",
                height=200,
            )
            self.clear_locations_container.pack(pady=(10, 0), fill="both", expand=False)

        # â”€â”€â”€ Audit routes storage â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # maps dispatch location â†’ set of routes that currently need auditing
        self.clear_audit_routes = {}
        
        # â”€â”€â”€ SoF+ load retry counter â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.sof_load_attempts = 0
        
        # SoF+ tab content is already deferred separately (see _build_sof_ui_once)

        # â”€â”€â”€ Auto-Sort Tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # sort_tab = tab.tab("Auto-Sort")  # Will be created dynamically when Shift+F6 is pressed
        # from customtkinter import CTkProgressBar  # Moved to _create_auto_sort_ui method

        # sort_frame = ctk.CTkFrame(  # Moved to _create_auto_sort_ui method
        #     sort_tab,
        #     fg_color="black",
        #     border_width=2,
        #     border_color=BRAND_ORANGE
        # )
        # sort_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        # # Top box: Controls and progress
        # top_box = ctk.CTkFrame(  # Moved to _create_auto_sort_ui method
        #     sort_frame,
        #     fg_color="black",
        #     border_width=2,
        #     border_color=BRAND_ORANGE
        # )
        # top_box.pack(padx=10, pady=(10,5), fill="x")

        # # Centered header
        # ctk.CTkLabel(  # Moved to _create_auto_sort_ui method
        #     top_box,
        #     text="Sort Regular Pickups",
        #     font=ctk.CTkFont(size=20, weight="bold"),
        #     text_color=BRAND_ORANGE
        # ).pack(pady=(10,15))

        # # Controls row: Location entry + Start/Stop buttons
        # ctrl = ctk.CTkFrame(top_box, fg_color="black")  # Moved to _create_auto_sort_ui method
        # ctrl.pack(padx=10, pady=(0,10))

        # ctk.CTkLabel(  # Moved to _create_auto_sort_ui method
        #     ctrl,
        #     text="Location:",
        #     text_color=BRAND_ORANGE,
        #     font=ctk.CTkFont(size=16, weight="bold")
        # ).pack(side="left", padx=(0,5))

        # self.sort_location_entry = ctk.CTkEntry(  # Moved to _create_auto_sort_ui method
        #     ctrl,
        #     width=100,
        #     fg_color="white",     # white background
        #     text_color="black"    # black text
        # )
        # self.sort_location_entry.pack(side="left")
        # # â—€ Bind to uppercase + max-4 helper
        # self.sort_location_entry.bind(
        #     "<KeyRelease>",
        #     lambda e: self._convert_to_uppercase_sort()
        # )
        # # â—€ Bind Enter key to start sort function
        # self.sort_location_entry.bind(
        #     "<Return>",
        #     lambda e: self.start_sort()
        # )

        # self.sort_start_button = ctk.CTkButton(  # Moved to _create_auto_sort_ui method
        #     ctrl,
        #     text="Start",
        #     fg_color=BRAND_ORANGE,
        #     hover_color=BRAND_PURPLE,
        #     command=self.start_sort
        # )
        # self.sort_start_button.pack(side="left", padx=(5,5))

        # self.sort_stop_button = ctk.CTkButton(  # Moved to _create_auto_sort_ui method
        #     ctrl,
        #     text="Stop",
        #     fg_color="red",
        #     hover_color="darkred",
        #     command=self.stop_sort,
        #     state="disabled"  # Initially disabled until process starts
        # )
        # self.sort_stop_button.pack(side="left", padx=(0,0))

        # # Custom progress bar with Carrier truck and road
        # self.progress_frame = ctk.CTkFrame(top_box, fg_color="black", height=60)  # Moved to _create_auto_sort_ui method
        # self.progress_frame.pack(padx=10, pady=(10,5), fill="x")
        # self.progress_frame.pack_propagate(False)
        
        # # Road background (black with orange lines)
        # self.road_canvas = tk.Canvas(  # Moved to _create_auto_sort_ui method
        #     self.progress_frame,
        #     bg="black",  # Black road to match UI
        #     height=40,
        #     highlightthickness=0
        # )
        # self.road_canvas.pack(pady=10, fill="x", padx=5)
        
        # # Create progress overlay canvas
        # self.progress_canvas = tk.Canvas(  # Moved to _create_auto_sort_ui method
        #     self.progress_frame,
        #     bg="black",
        #     height=40,
        #     highlightthickness=0
        # )
        # self.progress_canvas.place(x=5, y=10, relwidth=0, height=40)
        
        # # Draw road lines (initially empty)
        # self.draw_road_lines(0)  # Moved to _create_auto_sort_ui method
        
        # # Load Carrier truck logo  # Moved to _create_auto_sort_ui method
        # try:
        #     from PIL import Image, ImageTk
        #     truck_img = Image.open("d:/Portable/scripts/TRUCK_LOGO.png")
        #     # Resize to fit progress bar height
        #     truck_img = truck_img.resize((50, 30), Image.LANCZOS)
        #     self.truck_photo = ImageTk.PhotoImage(truck_img)
        #     
        #     # Create truck on main road canvas (initially hidden)
        #     self.truck_id = self.road_canvas.create_image(
        #         10, 20,  # Start position
        #         image=self.truck_photo,
        #         anchor="w",
        #         state="hidden"  # Initially hidden
        #     )
        # except Exception as e:
        #     print(f"Could not load truck logo: {e}")
        #     # Fallback: create a simple truck shape
        #     self.truck_id = self.road_canvas.create_rectangle(
        #         10, 10, 60, 30,
        #         fill=BRAND_ORANGE,
        #         outline="white",
        #         state="hidden"  # Initially hidden
        #     )
        # 
        # # Truck breathing animation variables  # Moved to _create_auto_sort_ui method
        # self.truck_breathing_alpha = 1.0
        # self.truck_breathing_direction = -1
        # self.truck_breathing_active = False
        # 
        # # Status frame to hold status label and percentage side by side  # Moved to _create_auto_sort_ui method
        # status_frame = ctk.CTkFrame(top_box, fg_color="black")
        # status_frame.pack(pady=(5,10))
        # 
        # # Status label for running animation  # Moved to _create_auto_sort_ui method
        # self.sort_status_label = ctk.CTkLabel(
        #     status_frame,
        #     text="",
        #     text_color=BRAND_ORANGE,
        #     font=ctk.CTkFont(size=14, weight="bold")
        # )
        # self.sort_status_label.pack(side="left", padx=(10,5))
        # 
        # # Progress text positioned next to status label  # Moved to _create_auto_sort_ui method
        # self.progress_text = ctk.CTkLabel(
        #     status_frame,
        #     text="",
        #     text_color=BRAND_ORANGE,
        #     font=ctk.CTkFont(size=12, weight="bold")
        # )
        # self.progress_text.pack(side="left", padx=(5,10))
        # self.sort_breathing_alpha = 1.0
        # self.sort_breathing_direction = -1

        # # Bottom box: Results (only shown when mismatches found)  # Moved to _create_auto_sort_ui method
        # self.bottom_box = ctk.CTkFrame(
        #     sort_frame,
        #     fg_color="black",
        #     border_width=2,
        #     border_color=BRAND_ORANGE
        # )
        # # Don't pack initially - only pack when results are found
        # 
        # # Results header with copy button  # Moved to _create_auto_sort_ui method
        # header_frame = ctk.CTkFrame(self.bottom_box, fg_color="black")
        # header_frame.pack(fill="x", padx=10, pady=(10,5))
        # 
        # self.results_header = ctk.CTkLabel(
        #     header_frame,
        #     text="Stops Mismatched/Out of place:",
        #     text_color=BRAND_ORANGE,
        #     font=ctk.CTkFont(size=14, weight="bold")
        # )
        # self.results_header.pack(side="left", padx=(0,10))
        # 
        # # Copy results button  # Moved to _create_auto_sort_ui method
        # self.copy_results_button = ctk.CTkButton(
        #     header_frame,
        #     text="ðŸ“‹ Copy Results",
        #     width=120,
        #     height=25,
        #     font=ctk.CTkFont(size=12, weight="bold"),
        #     fg_color=BRAND_ORANGE,
        #     hover_color=BRAND_PURPLE,
        #     command=self.copy_results_to_clipboard
        # )
        # self.copy_results_button.pack(side="right", padx=(0,5))
        # 
        # # Auto Move button (initially invisible until shift-f6 is pressed)  # Moved to _create_auto_sort_ui method
        # self.stop_movement_button = ctk.CTkButton(
        #     header_frame,
        #     text="Auto Move",
        #     width=100,
        #     height=25,
        #     font=ctk.CTkFont(size=12, weight="bold"),
        #     fg_color="gray",
        #     hover_color="gray",
        #     state="disabled",
        #     command=self.start_stop_movement
        # )
        # # Don't pack initially - will be packed when Shift+F6 is pressed
        

        
        # # Container for result labels  # Moved to _create_auto_sort_ui method
        # self.results_container = ctk.CTkScrollableFrame(self.bottom_box, fg_color="black", width=580, height=250)
        # self.results_container.pack(padx=10, pady=(0,10), fill="both", expand=True)
        # 
        # # Select all checkbox at the top (will be hidden during auto move)  # Moved to _create_auto_sort_ui method
        # self.select_all_var = tk.BooleanVar()
        # self.select_all_checkbox = ctk.CTkCheckBox(
        #     self.results_container,
        #     text="Select All",
        #     variable=self.select_all_var,
        #     command=self.toggle_select_all,
        #     fg_color=BRAND_ORANGE,
        #     hover_color=BRAND_PURPLE,
        #     text_color="white",
        #     font=ctk.CTkFont(size=12, weight="bold")
        # )
        # self.select_all_checkbox.pack(anchor="w", padx=5, pady=(5,10))
        # 
        # # List to store result frames (checkbox + label)  # Moved to _create_auto_sort_ui method
        # self.result_frames = []

        # --- SoF+ Tab ---
        sof = self.tab.tab("SoF+")

        # Defer heavy SoF+ UI construction until the tab is opened
        self._sof_ui_initialized = getattr(self, "_sof_ui_initialized", False)

        def _build_sof_ui_once():
            if self._sof_ui_initialized:
                return
            self._sof_ui_initialized = True

            # 1) Centered Entry + Load button
            entry_frame = ctk.CTkFrame(sof, fg_color="black")
            entry_frame.pack(pady=(20,10))                # no fillâ€”so it centers
            entry_frame.grid_columnconfigure(0, weight=1) # allow centering

            self.sof_location_entry = ctk.CTkEntry(
                entry_frame,
                placeholder_text="Dispatch Location",
                width=60,
                fg_color="white", text_color="black", border_color=BRAND_ORANGE
            )
            self.sof_location_entry.grid(row=0, column=0, padx=(0,5))
            self.sof_location_entry.bind("<Return>", lambda e: self.load_sof_routes())
            # Limit to 4 uppercase chars:
            self.sof_location_entry.bind(
                "<KeyRelease>",
                lambda e: self._convert_to_uppercase_sof()
            )

            self.sof_load_button = ctk.CTkButton(
                entry_frame, text="Load",
                width=60, height=25,
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                command=self.load_sof_routes
            )
            self.sof_load_button.grid(row=0, column=1, padx=(5,0))

            # â”€â”€â”€ Reset button (disabled until routes are loaded) â”€â”€â”€â”€â”€â”€â”€
            self.sof_reset_button = ctk.CTkButton(
                entry_frame, text="Reset",
                width=60, height=25,
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                state="disabled",
                command=self.reset_sof_tab
            )
            self.sof_reset_button.grid(row=0, column=2, padx=(5,0))

            # â†“ half-sized Legend button â†“
            self.sof_legend_button = ctk.CTkButton(
                entry_frame, text="Legend",
                width=60, height=25,
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                command=self.show_legend_popup
            )
            self.sof_legend_button.grid(row=0, column=3, padx=(5,0))

            # 1b) Statistics counters
            stats_frame = ctk.CTkFrame(sof, fg_color="black")
            stats_frame.pack(pady=(10, 5), fill="x", padx=10)
            
            # Configure 7 grid columns to be evenly spaced (5 main + 2 secondary)
            for col in range(7):
                stats_frame.grid_columnconfigure(col, weight=1)
            
            # Define consistent font for all counter labels
            counter_label_font = ctk.CTkFont(size=16, weight="bold")
            
            # UNA counter
            una_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            una_frame.grid(row=0, column=0, padx=10, pady=5)
            self.una_label = ctk.CTkLabel(
                una_frame, text="UNA",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.una_counter = AnimatedSpinnerCounter(una_frame, size=30, color="#FF6600")
            self.una_counter.pack()
            # Click anywhere on frame/number/label to open UNA details
            una_click = lambda e: self.show_una_popup()
            una_frame.bind("<Button-1>", una_click)
            self.una_label.bind("<Button-1>", una_click)
            try:
                self.una_counter.frame.bind("<Button-1>", una_click)
                self.una_counter.canvas.bind("<Button-1>", una_click)
                if self.una_counter.text_label:
                    self.una_counter.text_label.bind("<Button-1>", una_click)
            except Exception:
                pass
            # Tooltip for UNA (UNASSIGNED STOPS)
            try:
                CustomTooltip(self.una_label, "UNASSIGNED STOPS")
            except Exception:
                pass

            # EARLY counter
            early_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            early_frame.grid(row=0, column=1, padx=10, pady=5)
            self.early_label = ctk.CTkLabel(
                early_frame, text="EARLY",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.early_counter = AnimatedSpinnerCounter(early_frame, size=30, color="#FF6600")
            self.early_counter.pack()
            early_click = lambda e: self.show_early_popup()
            early_frame.bind("<Button-1>", early_click)
            self.early_label.bind("<Button-1>", early_click)
            try:
                self.early_counter.frame.bind("<Button-1>", early_click)
                self.early_counter.canvas.bind("<Button-1>", early_click)
                if self.early_counter.text_label:
                    self.early_counter.text_label.bind("<Button-1>", early_click)
            except Exception:
                pass
            # Tooltip for EARLY (EARLY PICK UPS)
            try:
                CustomTooltip(self.early_label, "EARLY PICK UPS")
            except Exception:
                pass

            # LATE counter
            late_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            late_frame.grid(row=0, column=2, padx=10, pady=5)
            self.late_label = ctk.CTkLabel(
                late_frame, text="LATE",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.late_counter = AnimatedSpinnerCounter(late_frame, size=30, color="#FF6600")
            self.late_counter.pack()
            late_click = lambda e: self.show_late_popup()
            late_frame.bind("<Button-1>", late_click)
            self.late_label.bind("<Button-1>", late_click)
            try:
                self.late_counter.frame.bind("<Button-1>", late_click)
                self.late_counter.canvas.bind("<Button-1>", late_click)
                if self.late_counter.text_label:
                    self.late_counter.text_label.bind("<Button-1>", late_click)
            except Exception:
                pass
            # Tooltip for LATE (LATE PICK UPS)
            try:
                CustomTooltip(self.late_label, "LATE PICK UPS")
            except Exception:
                pass

            # UNTRANSMITTED counter (displayed as UNT with tooltip)
            untransmitted_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            untransmitted_frame.grid(row=0, column=3, padx=10, pady=5)
            self.untransmitted_label = ctk.CTkLabel(
                untransmitted_frame, text="UNT",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.untransmitted_counter = AnimatedSpinnerCounter(untransmitted_frame, size=30, color="#FF6600")
            self.untransmitted_counter.pack()
            unt_click = lambda e: self.show_untransmitted_popup()
            untransmitted_frame.bind("<Button-1>", unt_click)
            self.untransmitted_label.bind("<Button-1>", unt_click)
            try:
                self.untransmitted_counter.frame.bind("<Button-1>", unt_click)
                self.untransmitted_counter.canvas.bind("<Button-1>", unt_click)
                if self.untransmitted_counter.text_label:
                    self.untransmitted_counter.text_label.bind("<Button-1>", unt_click)
            except Exception:
                pass
            # Tooltip for UNT (UNTRANSMITTED STOPS)
            try:
                CustomTooltip(self.untransmitted_label, "UNTRANSMITTED STOPS")
            except Exception:
                pass

            # DUPLICATES counter (labelled DUPS)
            duplicates_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            duplicates_frame.grid(row=0, column=4, padx=10, pady=5)
            self.duplicates_label = ctk.CTkLabel(
                duplicates_frame, text="DUPS",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.duplicates_counter = AnimatedSpinnerCounter(duplicates_frame, size=30, color="#FF6600")
            self.duplicates_counter.pack()
            dups_click = lambda e: self.show_duplicates_popup()
            duplicates_frame.bind("<Button-1>", dups_click)
            self.duplicates_label.bind("<Button-1>", dups_click)
            try:
                self.duplicates_counter.frame.bind("<Button-1>", dups_click)
                self.duplicates_counter.canvas.bind("<Button-1>", dups_click)
                if self.duplicates_counter.text_label:
                    self.duplicates_counter.text_label.bind("<Button-1>", dups_click)
            except Exception:
                pass
            # Tooltip for DUPS (DUPLICATE STOPS)
            try:
                CustomTooltip(self.duplicates_label, "DUPLICATE STOPS")
            except Exception:
                pass

            # DEX 03 counter (bad addresses) - only visible after routes loaded
            self.dex03_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            self.dex03_frame.grid(row=0, column=5, padx=10, pady=5)
            self.dex03_label = ctk.CTkLabel(
                self.dex03_frame, text="DEX 03",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.dex03_counter = AnimatedSpinnerCounter(self.dex03_frame, size=30, color="#FF6600")
            self.dex03_counter.pack()
            self.dex03_label.pack()
            self.dex03_counter.show_text("0")  # Placeholder
            self.dex03_frame.bind("<Button-1>", lambda e: getattr(self, "show_dex03_popup", lambda: None)())
            self.dex03_label.bind("<Button-1>", lambda e: getattr(self, "show_dex03_popup", lambda: None)())
            try:
                CustomTooltip(self.dex03_label, "BAD ADDRESSES")
            except Exception:
                pass
            # Hide initially until routes are loaded
            self.dex03_frame.grid_remove()

            # PUX 03 counter (bad pickup address) - only visible after routes loaded
            self.pux03_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            self.pux03_frame.grid(row=0, column=6, padx=10, pady=5)
            self.pux03_label = ctk.CTkLabel(
                self.pux03_frame, text="PUX 03",
                font=counter_label_font,
                text_color="#FF6600"
            )
            self.pux03_counter = AnimatedSpinnerCounter(self.pux03_frame, size=30, color="#FF6600")
            self.pux03_counter.pack()
            self.pux03_label.pack()
            self.pux03_counter.show_text("0")  # Placeholder
            self.pux03_frame.bind("<Button-1>", lambda e: getattr(self, "show_pux03_popup", lambda: None)())
            self.pux03_label.bind("<Button-1>", lambda e: getattr(self, "show_pux03_popup", lambda: None)())
            try:
                CustomTooltip(self.pux03_label, "BAD PICKUP ADDRESS")
            except Exception:
                pass
            # Hide initially until routes are loaded
            self.pux03_frame.grid_remove()

            # Initialize counters and popup data
            self.una_count = 0
            self.early_count = 0
            self.late_count = 0
            self.untransmitted_count = 0
            self.duplicates_count = 0
            self.una_popup_data = []
            self.early_popup_data = []
            self.late_popup_data = []
            self.untransmitted_popup_data = {}
            self.duplicates_popup_data = {}
            self.organized_duplicates_data = {'dups': {}, 'non_dups': {}}
            self.organized_untransmitted_data = {'unsafe': {}, 'safe': {}}
            self.removed_duplicates = set()

            # Hide labels initially
            self.una_label.pack_forget()
            self.early_label.pack_forget()
            self.late_label.pack_forget()
            self.untransmitted_label.pack_forget()
            self.duplicates_label.pack_forget()

            # Test function scheduling and monitor state
            self.test_scheduler_active = False
            self.test_scheduler_stop_event = threading.Event()
            self.sof_monitor_thread = None
            self.sof_monitor_stop_event = threading.Event()

            # 1c) Loading bars (hidden until needed)
            from customtkinter import CTkProgressBar
            self.sof_progress = CTkProgressBar(sof, orientation="horizontal", fg_color="black", progress_color=BRAND_ORANGE)
            self.sof_progress.set(0)
            self.sof_monitor_progress = CTkProgressBar(sof, orientation="horizontal", fg_color="black", progress_color=BRAND_ORANGE)
            self.sof_monitor_progress.set(0)

            # 1d) SoF+ routes grid (hidden until load)
            self.sof_grid_frame = ctk.CTkScrollableFrame(sof, fg_color="black", width=580, height=400)
            # Speed up mouse-wheel scrolling inside the SoF+ grid
            self._enable_fast_scroll(self.sof_grid_frame, lines=6)

            # 1e) Status label
            self.test_status_label = ctk.CTkLabel(sof, text="", font=ctk.CTkFont(size=12), text_color="white")
            self.test_status_label.pack(pady=(5, 0))

        # Defer build shortly after UI loads to keep login snappy without breaking tabs
        self.after(0, _build_sof_ui_once)

        # --- Reports Tab ---
        reports_tab = self.tab.tab("Reports")

        # OPS Package Count UI, centered within the Reports tab
        ops_frame = ctk.CTkFrame(reports_tab, fg_color="black")
        ops_frame.pack(pady=(60, 20))
        for col in range(3):
            ops_frame.grid_columnconfigure(col, weight=1)

        header_label = ctk.CTkLabel(
            ops_frame,
            text="OPS Package Count",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=BRAND_ORANGE,
        )
        header_label.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky="n")

        # Scan Date row
        date_label = ctk.CTkLabel(
            ops_frame,
            text="Scan Date",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="white",
        )
        date_label.grid(row=1, column=0, padx=(0, 5), pady=(0, 10), sticky="e")

        self.ops_date_entry = ctk.CTkEntry(
            ops_frame,
            placeholder_text="MM/DD/YYYY",
            width=140,
            fg_color="white",
            text_color="black",
            border_color=BRAND_ORANGE,
        )
        self.ops_date_entry.grid(row=1, column=1, pady=(0, 10), sticky="w")

        date_btn = ctk.CTkButton(
            ops_frame,
            text="ðŸ“…",
            width=30,
            height=28,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self._open_ops_date_picker,
        )
        date_btn.grid(row=1, column=2, padx=(5, 0), pady=(0, 10), sticky="w")

        # Destination Location row
        dest_label = ctk.CTkLabel(
            ops_frame,
            text="Destination Location",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="white",
        )
        dest_label.grid(row=2, column=0, padx=(0, 5), pady=(0, 10), sticky="e")

        self.ops_location_entry = ctk.CTkEntry(
            ops_frame,
            placeholder_text="e.g. HWOA",
            width=140,
            fg_color="white",
            text_color="black",
            border_color=BRAND_ORANGE,
        )
        self.ops_location_entry.grid(row=2, column=1, pady=(0, 10), sticky="w")
        self.ops_location_entry.bind("<KeyRelease>", self._convert_to_uppercase_ops_location)

        # Run button (centered under the fields)
        self.ops_run_button = ctk.CTkButton(
            ops_frame,
            text="Run OPS Package Count",
            width=220,
            height=34,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self.run_ops_package_count_threaded,
        )
        self.ops_run_button.grid(row=3, column=0, columnspan=3, pady=(18, 0))

        # Status label under the controls
        self.ops_status_label = ctk.CTkLabel(
            reports_tab,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="white",
        )
        self.ops_status_label.pack(pady=(10, 0))

        # --- Settings Tab ---
        settings_tab = self.tab.tab("Settings")
        ctk.CTkLabel(settings_tab, text="Coming Soon",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=BRAND_ORANGE).pack(pady=(80, 10))
        
        # Update widget button visibility based on current tab
        self.update_widget_button_visibility()

    def _format_ops_calendar_aria(self, date_str: str) -> str:
        """
        Convert a user-entered date into the aria-label format used
        by the Angular Material calendar, e.g. 'March 12, 2026'.

        Accepts either MM/DD/YYYY or YYYY-MM-DD.
        """
        parsed = None
        for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(date_str, fmt)
                break
            except ValueError:
                continue
        if not parsed:
            raise ValueError("Enter date as MM/DD/YYYY or YYYY-MM-DD.")
        month_name = parsed.strftime("%B")
        # Avoid zero-padded day; Angular Material uses 'March 3, 2026' not 'March 03, 2026'
        day = parsed.day
        return f"{month_name} {day}, {parsed.year}"

    def _open_ops_date_picker(self):
        """Open a simple calendar popup to pick the OPS scan date."""
        # If a picker is already open, don't open another
        if getattr(self, "_ops_date_popup", None) and self._ops_date_popup.winfo_exists():
            self._ops_date_popup.lift()
            return

        now = datetime.now()
        year = now.year
        month = now.month

        # Try to parse currently entered date to start the picker there
        current = self.ops_date_entry.get().strip()
        for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                d = datetime.strptime(current, fmt)
                year, month = d.year, d.month
                break
            except ValueError:
                continue

        self._build_ops_calendar_popup(year, month)

    def _build_ops_calendar_popup(self, year: int, month: int):
        import calendar as _cal

        # Destroy any existing popup
        if getattr(self, "_ops_date_popup", None) and self._ops_date_popup.winfo_exists():
            self._ops_date_popup.destroy()

        popup = tk.Toplevel(self)
        # Borderless popup without OS title bar/tab
        try:
            popup.overrideredirect(True)
        except Exception:
            pass
        popup.configure(bg="black")
        popup.resizable(False, False)
        try:
            # Keep the calendar above the main ROUTR window while it is open
            top = self.winfo_toplevel()
            popup.transient(top)
            popup.lift()
            popup.attributes("-topmost", True)
        except Exception:
            pass
        self._ops_date_popup = popup

        # Position the popup directly under the OPS date entry, so it is fully visible
        try:
            self.update_idletasks()
            ex = self.ops_date_entry.winfo_rootx()
            ey = self.ops_date_entry.winfo_rooty()
            eh = self.ops_date_entry.winfo_height()
            popup.update_idletasks()
            pw = popup.winfo_reqwidth()
            ph = popup.winfo_reqheight()

            # Default position just under the entry
            x = ex
            y = ey + eh + 4

            # Keep within screen bounds if possible
            screen_w = popup.winfo_screenwidth()
            screen_h = popup.winfo_screenheight()
            if x + pw > screen_w:
                x = max(0, screen_w - pw - 10)
            if y + ph > screen_h:
                # If it would go off the bottom, show it above the entry instead
                y = max(0, ey - ph - 4)

            popup.geometry(f"+{x}+{y}")
        except Exception:
            # Fallback: center in main window if anything goes wrong
            self._position_popup(popup)

        header_frame = tk.Frame(popup, bg="black")
        header_frame.pack(pady=(6, 4), padx=6)

        month_name = _cal.month_name[month]
        lbl = tk.Label(
            header_frame,
            text=f"{month_name} {year}",
            bg="black",
            fg=BRAND_ORANGE,
            font=("Arial", 11, "bold"),
        )
        lbl.grid(row=0, column=1, padx=10)

        def prev_month():
            new_month = month - 1
            new_year = year
            if new_month == 0:
                new_month = 12
                new_year -= 1
            self._build_ops_calendar_popup(new_year, new_month)

        def next_month():
            new_month = month + 1
            new_year = year
            if new_month == 13:
                new_month = 1
                new_year += 1
            self._build_ops_calendar_popup(new_year, new_month)

        tk.Button(
            header_frame,
            text="<",
            command=prev_month,
            bg="black",
            fg="white",
            relief="flat",
        ).grid(row=0, column=0)
        tk.Button(
            header_frame,
            text=">",
            command=next_month,
            bg="black",
            fg="white",
            relief="flat",
        ).grid(row=0, column=2)

        body = tk.Frame(popup, bg="black")
        body.pack(padx=8, pady=(0, 10))

        # Weekday headers (Sunday-first to match the calendar grid)
        for i, wd in enumerate(["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]):
            tk.Label(body, text=wd, bg="black", fg="white").grid(row=0, column=i, padx=3, pady=2)

        # Use Sunday as the first weekday so our headers align correctly
        cal = _cal.Calendar(firstweekday=6)
        row = 1
        for week in cal.monthdayscalendar(year, month):
            for col, day in enumerate(week):
                if day == 0:
                    tk.Label(body, text=" ", bg="black", fg="white", width=3).grid(
                        row=row, column=col, padx=1, pady=1
                    )
                else:
                    def make_cmd(d=day, m=month, y=year):
                        return lambda: self._on_ops_day_selected(y, m, d)

                    tk.Button(
                        body,
                        text=str(day),
                        width=4,
                        command=make_cmd(),
                        bg="#222222",
                        fg="white",
                        relief="flat",
                    ).grid(row=row, column=col, padx=1, pady=1)
            row += 1

    def _on_ops_day_selected(self, year: int, month: int, day: int):
        """Handle clicking a date in the OPS calendar popup."""
        picked = datetime(year, month, day)
        self.ops_date_entry.delete(0, ctk.END)
        self.ops_date_entry.insert(0, picked.strftime("%m/%d/%Y"))

        if getattr(self, "_ops_date_popup", None) and self._ops_date_popup.winfo_exists():
            self._ops_date_popup.destroy()

    def run_ops_package_count_threaded(self):
        """Kick off OPS Package Count in a background thread."""
        date_str = self.ops_date_entry.get().strip()
        dest_loc = self.ops_location_entry.get().strip().upper()

        if not date_str or not dest_loc:
            messagebox.showerror("Missing Info", "Enter both Scan Date and Destination Location.")
            return

        self.ops_run_button.configure(state="disabled")
        self.ops_status_label.configure(text="Launching OPS Package Count browserâ€¦")

        t = threading.Thread(
            target=self._run_ops_package_count_worker,
            args=(date_str, dest_loc),
            daemon=True,
        )
        t.start()

    def _run_ops_package_count_worker(self, date_str: str, dest_loc: str):
        """
        OPS Package Count via API: log in to OPS UI, capture Bearer token,
        then call package-tracking-exception-list API with requests.
        """
        driver = None
        try:
            logging.info("OPS: worker starting with date=%s dest_loc=%s", date_str, dest_loc)
            # Parse date to YYYY-MM-DD for API
            try:
                parsed = datetime.strptime(date_str.strip(), "%m/%d/%Y")
            except ValueError:
                try:
                    parsed = datetime.strptime(date_str.strip(), "%Y-%m-%d")
                except ValueError:
                    self.after(
                        0,
                        lambda: messagebox.showerror("Invalid Date", "Enter date as MM/DD/YYYY or YYYY-MM-DD."),
                    )
                    self.after(0, lambda: self.ops_run_button.configure(state="normal"))
                    self.after(0, lambda: self.ops_status_label.configure(text=""))
                    return
            api_date = parsed.strftime("%Y-%m-%d")
            station = dest_loc.strip().lower()
            carrierx_id = (self.username or "").strip()
            if not carrierx_id:
                self.after(
                    0,
                    lambda: messagebox.showerror("Not Logged In", "Log in first so Carrier ID is available for the API."),
                )
                self.after(0, lambda: self.ops_run_button.configure(state="normal"))
                self.after(0, lambda: self.ops_status_label.configure(text=""))
                return

            self.after(0, lambda: self.ops_status_label.configure(text="Opening OPS browserâ€¦"))
            driver_path = ChromeDriverManager().install()
            service = Service(driver_path)
            opts = Options()
            opts.add_argument("--disable-extensions")
            opts.add_argument("--start-maximized")
            # Run headless so this browser is not visible to the user
            try:
                opts.add_argument("--headless=new")
            except Exception:
                opts.add_argument("--headless")
            opts.set_capability("goog:loggingPrefs", {"performance": "ALL"})
            driver = webdriver.Chrome(service=service, options=opts)
            wait = WebDriverWait(driver, 20)

            # Navigate to the legacy OPS report launcher twice on eops-lb-las, as you do manually.
            # 1) First hit to wake up the session and let redirects happen.
            ops_launcher_url = (
                "https://internal.example.com"
                "eShipmentGUI/DisplayLinkHandler?sublinkId=451"
            )
            self.after(0, lambda: self.ops_status_label.configure(text="Opening OPS launcherâ€¦"))
            driver.get(ops_launcher_url)
            try:
                WebDriverWait(driver, 30).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
            except Exception:
                logging.warning("OPS: first load of DisplayLinkHandler did not reach readyState=complete in time")
            time.sleep(1)

            # 2) Second hit to land on the actual Package Tracking Exception List page.
            self.after(0, lambda: self.ops_status_label.configure(text="Opening OPS report pageâ€¦"))
            driver.get(ops_launcher_url)
            try:
                WebDriverWait(driver, 30).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
            except Exception:
                logging.warning("OPS: second load of DisplayLinkHandler did not reach readyState=complete in time")
            time.sleep(2)
            logging.info("OPS: on OPS exception list page, current_url=%s", driver.current_url)

            self.after(0, lambda: self.ops_status_label.configure(text="Getting tokenâ€¦"))

            bearer_token = None
            # Try to read Bearer token from localStorage / sessionStorage (SSO and common keys)
            get_token_script = """
            (function(){
              function fromStorage(storage) {
                if (!storage) return null;
                // First try common SSO/token keys
                var keys = ['sso-token-storage', 'sso-cache-storage', 'accessToken', 'token', 'auth_token'];
                function extractToken(data) {
                  if (!data) return null;
                  try {
                    if (typeof data === 'string') {
                      // JWT-ish token
                      if (data.split('.').length >= 3 && data.length > 100) return data;
                      return null;
                    }
                    // SSO token manager shapes
                    if (data.accessToken) {
                      var at = data.accessToken.accessToken || data.accessToken.value || data.accessToken;
                      if (typeof at === 'string' && at.split('.').length >= 3 && at.length > 100) return at;
                    }
                    if (data.tokens && data.tokens.accessToken) {
                      var at2 = data.tokens.accessToken.accessToken || data.tokens.accessToken.value || data.tokens.accessToken;
                      if (typeof at2 === 'string' && at2.split('.').length >= 3 && at2.length > 100) return at2;
                    }
                    if (data.token && typeof data.token === 'string') return data.token;
                    if (data.access_token && typeof data.access_token === 'string') return data.access_token;
                  } catch(e) {}
                  return null;
                }
                for (var i = 0; i < keys.length; i++) {
                  try {
                    var raw = storage.getItem(keys[i]);
                    if (!raw) continue;
                    var parsed = raw;
                    try { parsed = JSON.parse(raw); } catch(e) {}
                    var t = extractToken(parsed);
                    if (t) return t;
                  } catch(e) {}
                }
                // Fallback: scan all keys for any JWT-like token
                try {
                  for (var j = 0; j < storage.length; j++) {
                    var k = storage.key(j);
                    if (!k) continue;
                    var v = storage.getItem(k);
                    if (!v) continue;
                    // quick check for JWT header prefix
                    if (v.indexOf('eyJ') === -1) continue;
                    var parsed2 = v;
                    try { parsed2 = JSON.parse(v); } catch(e) {}
                    var t2 = extractToken(parsed2);
                    if (t2) return t2;
                  }
                } catch(e) {}
                return null;
              }
              return fromStorage(window.localStorage) || fromStorage(window.sessionStorage);
            })();
            """
            # If we're redirected to SSO/Identity Portal, wait for the user to finish login
            try:
                start = time.time()
                while True:
                    cur = (driver.current_url or "").lower()
                    if "sso.com" not in cur and "identity" not in cur and "authorize" not in cur:
                        break
                    if time.time() - start > 180:
                        break
                    self.after(0, lambda: self.ops_status_label.configure(
                        text="Complete SSO login in the opened browserâ€¦"
                    ))
                    time.sleep(2)
                logging.info("OPS: after login wait current_url=%s", driver.current_url)
                # After login completes (we should be on MenuPage.iface), hit the OPS
                # DisplayLinkHandler twice exactly as you do manually, so we land on
                # the real Package Tracking Exception List page.
                for idx in range(2):
                    try:
                        self.after(0, lambda i=idx: self.ops_status_label.configure(
                            text="Opening OPS launcher (step {} of 2)â€¦".format(i + 1)
                        ))
                        driver.get(ops_launcher_url)
                        WebDriverWait(driver, 30).until(
                            lambda d: d.execute_script("return document.readyState") == "complete"
                        )
                        time.sleep(2 if idx == 1 else 1)
                    except Exception as e_nav:
                        logging.warning("OPS: error navigating DisplayLinkHandler step %s: %s", idx + 1, e_nav)
                        break
                logging.info("OPS: after DisplayLinkHandler hits, current_url=%s", driver.current_url)
            except Exception:
                pass
            try:
                bearer_token = driver.execute_script(get_token_script)
            except Exception as e:
                logging.warning("OPS: could not read token from storage: %s", e)

            # Fallback: capture Authorization header from performance log (any request that has it)
            if not bearer_token:
                try:
                    for entry in driver.get_log("performance") or []:
                        try:
                            msg = json.loads(entry.get("message", "{}"))
                            m = msg.get("message", {})
                            if m.get("method") != "Network.requestWillBeSent":
                                continue
                            req = (m.get("params") or {}).get("request") or {}
                            url = (req.get("url") or "")
                            if "package-tracking-exception-list" in url or "internal-api" in url:
                                headers = req.get("headers") or {}
                                auth = headers.get("Authorization") or headers.get("authorization")
                                if auth and "Bearer " in auth:
                                    bearer_token = auth.replace("Bearer ", "").strip()
                                    break
                        except (json.JSONDecodeError, KeyError, TypeError):
                            continue
                except Exception as e:
                    logging.warning("OPS: could not get token from performance log: %s", e)

            if not bearer_token:
                self.after(
                    0,
                    lambda: self.ops_status_label.configure(
                        text="Could not get OPS token from session. Try again in a moment."
                    ),
                )
                self.after(0, lambda: self.ops_run_button.configure(state="normal"))
                return

            self.after(0, lambda: self.ops_status_label.configure(text="Calling OPS APIâ€¦"))

            # Call API with requests (give it plenty of time; service can be slow)
            rows = get_ops_exception_list(
                station=station,
                date_str=api_date,
                track_type="STAT",
                exception_code="13",
                carrierx_id=carrierx_id,
                bearer_token=bearer_token,
                timeout=90,
            )
            count = len(rows) if rows else 0
            logging.info("OPS: API returned %s rows for %s on %s", count, station, api_date)

            # Build per-origin counts like "Origin: HWOKN - 45"
            origin_counts = {}
            for r in rows or []:
                origin = (r.get("origLocationCd") or "").strip()
                if not origin:
                    loc_info = (r.get("locationInfo") or "")
                    # e.g. "Origin: HWOKNDelivery: null"
                    if "Origin:" in loc_info:
                        try:
                            after = loc_info.split("Origin:", 1)[1].strip()
                            # stop at first space or "Delivery"
                            for sep in ["Delivery", " "]:
                                if sep in after:
                                    after = after.split(sep, 1)[0]
                                    break
                            origin = after.strip()
                        except Exception:
                            origin = ""
                if not origin:
                    origin = "UNKNOWN"
                origin_counts[origin] = origin_counts.get(origin, 0) + 1

            # Build time-segment (hour) + origin counts using local scan timestamp
            # trackTmstpLocalTm looks like "03/12/2026 16:01"
            time_origin_counts = {}
            for r in rows or []:
                ts = (r.get("trackTmstpLocalTm") or "").strip()
                if not ts:
                    continue
                try:
                    dt_local = datetime.strptime(ts, "%m/%d/%Y %H:%M")
                except Exception:
                    continue
                # Hour-level bucket, e.g. "16:00"
                bucket = dt_local.strftime("%H:00")
                origin = (r.get("origLocationCd") or "").strip()
                if not origin:
                    loc_info = (r.get("locationInfo") or "")
                    if "Origin:" in loc_info:
                        try:
                            after = loc_info.split("Origin:", 1)[1].strip()
                            for sep in ["Delivery", " "]:
                                if sep in after:
                                    after = after.split(sep, 1)[0]
                                    break
                            origin = after.strip()
                        except Exception:
                            origin = ""
                if not origin:
                    origin = "UNKNOWN"
                key = (bucket, origin)
                time_origin_counts[key] = time_origin_counts.get(key, 0) + 1

            # Sort by descending count
            sorted_counts = sorted(origin_counts.items(), key=lambda kv: kv[1], reverse=True)
            summary_lines = [f"Origin: {orig} - {cnt}" for orig, cnt in sorted_counts]
            logging.info("OPS: per-origin counts:\n%s", "\n".join(summary_lines))

            # Build per-hour summary like "16:00 - HWOKN: 45, GHMA: 22"
            hour_buckets = {}
            for (bucket, origin), cnt in time_origin_counts.items():
                hour_buckets.setdefault(bucket, {})[origin] = cnt
            sorted_hours = sorted(hour_buckets.items())  # ascending by hour label
            time_lines = []
            for bucket, origins_map in sorted_hours:
                inner = sorted(origins_map.items(), key=lambda kv: kv[1], reverse=True)
                inner_str = ", ".join(f"{o}: {c}" for o, c in inner)
                time_lines.append(f"{bucket} - {inner_str}")
            if time_lines:
                logging.info("OPS: per-hour origin counts:\n%s", "\n".join(time_lines))

            # Cache a plain-text summary for the Copy button
            summary_text_parts = []
            summary_text_parts.append(f"OPS Package Count: {count} exceptions for {station.upper()} on {api_date}.")
            if summary_lines:
                summary_text_parts.append("")
                summary_text_parts.append("OPS exceptions by origin:")
                summary_text_parts.extend(summary_lines)
            if time_lines:
                summary_text_parts.append("")
                summary_text_parts.append("OPS exceptions by hour and origin:")
                summary_text_parts.extend(time_lines)
            self.ops_last_summary_text = "\n".join(summary_text_parts)

            def _finish_success():
                self.ops_status_label.configure(
                    text=f"OPS Package Count: {count} exceptions for {station.upper()} on {api_date}."
                )
                self.ops_run_button.configure(state="normal")
                try:
                    text = getattr(self, "ops_last_summary_text", "").strip()
                    if text:
                        self._show_ops_summary_popup(text)
                except Exception:
                    # If popup fails (e.g. during shutdown), ignore
                    pass

            self.after(0, _finish_success)

        except requests.RequestException as e:
            logging.exception("OPS: API request failed")
            err_msg = str(e)
            try:
                if getattr(e, "response", None) is not None and getattr(e.response, "text", None):
                    err_msg = (e.response.text or str(e))[:250]
            except Exception:
                pass
            def _finish_error(msg=err_msg):
                self.ops_status_label.configure(text=f"API error: {msg}")
                self.ops_run_button.configure(state="normal")
            self.after(0, _finish_error)
        except Exception as e:
            logging.exception("OPS: Error running OPS Package Count")
            def _finish_error(msg=str(e)):
                self.ops_status_label.configure(text=f"Error: {msg}")
                self.ops_run_button.configure(state="normal")
            self.after(0, _finish_error)

    def _convert_to_uppercase_ops_location(self, event=None):
        txt = self.ops_location_entry.get().upper()
        self.ops_location_entry.delete(0, ctk.END)
        self.ops_location_entry.insert(0, txt)

    def _show_ops_summary_popup(self, text: str):
        """Show OPS summary in a centered popup with a Copy button."""
        popup = tk.Toplevel(self)
        popup.title("OPS Package Count Summary")
        popup.configure(bg="black")
        try:
            popup.transient(self)
        except Exception:
            pass

        # Size and center roughly over the main window
        popup.geometry("520x520")
        try:
            self.update_idletasks()
            x = self.winfo_rootx() + (self.winfo_width() // 2) - 260
            y = self.winfo_rooty() + (self.winfo_height() // 2) - 260
            popup.geometry(f"+{x}+{y}")
        except Exception:
            pass

        frame = ctk.CTkFrame(popup, fg_color="black")
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        label = ctk.CTkLabel(
            frame,
            text="OPS Package Count Summary",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE,
        )
        label.pack(pady=(0, 8))

        # Scrollable text widget for the summary
        text_widget = tk.Text(
            frame,
            wrap="word",
            bg="black",
            fg="white",
            insertbackground="white",
            height=20,
        )
        text_widget.insert("1.0", text)
        text_widget.configure(state="disabled")
        text_widget.pack(fill="both", expand=True, padx=2, pady=(0, 8))

        btn_row = ctk.CTkFrame(frame, fg_color="black")
        btn_row.pack(pady=(4, 0))

        def do_copy():
            try:
                self.clipboard_clear()
                self.clipboard_append(text)
                self.update()
            except Exception as e:
                logging.exception("OPS: failed to copy from popup: %s", e)

        copy_btn = ctk.CTkButton(
            btn_row,
            text="Copy Results",
            width=120,
            fg_color=BRAND_PURPLE,
            hover_color=BRAND_ORANGE,
            command=do_copy,
        )
        copy_btn.pack(side="left", padx=(0, 8))

        close_btn = ctk.CTkButton(
            btn_row,
            text="Close",
            width=80,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=popup.destroy,
        )
        close_btn.pack(side="left")

    def convert_to_uppercase(self, event=None):
        txt = self.dispatch_entry.get().upper()[:4]
        self.dispatch_entry.delete(0, ctk.END)
        self.dispatch_entry.insert(0, txt)

    def convert_to_uppercase_clear(self, event=None):
        txt = self.clear_dispatch_entry.get().upper()[:4]
        self.clear_dispatch_entry.delete(0, ctk.END)
        self.clear_dispatch_entry.insert(0, txt)

    # --- Auto-Read methods ---
    def add_location(self, with_console=False):
        if with_console and not getattr(self, "console_allocated_read", False):
            import ctypes, sys, logging

            # 1) Allocate the new console & rebind stdout/stderr
            ctypes.windll.kernel32.AllocConsole()
            sys.stdout = open("CONOUT$", "w", buffering=1)
            sys.stderr = open("CONOUT$", "w", buffering=1)

            # 2) Tear down ALL existing logging handlers
            logging.shutdown()                      # close any open handlers
            for handler in logging.root.handlers[:]:
                logging.root.removeHandler(handler)

            # 3) Re initialize logging with exactly one fresh handler
            logging.basicConfig(
                level=logging.INFO,
                format="%(message)s",
                force=True         # requires Python â‰¥3.8; removes any residual handlers
            )

            logging.raiseExceptions = False
            logging.Handler.handleError = lambda self, record: None

            self.console_allocated_read = True
            logging.info("ðŸ” Console allocated for Auto-Read")

        dispatch = self.dispatch_entry.get().strip().upper()
        if not dispatch:
            messagebox.showerror("Missing Info", "Enter a dispatch location.")
            return
        if len(self.running_monitors) >= 5:
            messagebox.showerror("Limit Reached", "You may monitor up to 5 locations.")
            return
        if dispatch in self.running_monitors:
            messagebox.showerror("Duplicate", "Monitor already running for this location.")
            return
        # â”€â”€â”€ Use green-light check instead of text match â”€â”€â”€
        for m in self.running_monitors.values():
            fill = m["light"].itemcget(m["light"].oval_id, "fill")
            if fill != "#00FF00":  # not green
                messagebox.showerror("Busy", "Wait until existing monitors are idle.")
                return


        self.dispatch_entry.delete(0, ctk.END)
        types = []
        if self.pickup_reminder_var.get():
            types.append("Pickup Reminder")
        if self.early_pu_var.get():
            types.append("Early PU")
        if self.no_pickup_list_var.get():
            types.append("No Pickup List")
        if not types:
            messagebox.showerror("No Types", "Select at least one message type.")
            return

        stop_evt = threading.Event()
        frame = ctk.CTkFrame(self.locations_container, fg_color="black")
        frame.pack(fill="x", pady=5)
        ctk.CTkLabel(frame, text=f"Location: {dispatch}",
                     font=ctk.CTkFont(size=12), text_color="white").pack(side="left", padx=10)
        status_lbl = ctk.CTkLabel(frame, text="St: Startingâ€¦",
                                  font=ctk.CTkFont(size=12), text_color="white")
        status_lbl.pack(side="left", padx=10)
        light = ctk.CTkCanvas(frame, width=12, height=12, bg="black", highlightthickness=0)
        light.oval_id = light.create_oval(
            0, 0, 12, 12,
            fill=BRAND_ORANGE, outline=BRAND_ORANGE
        )
        light.pack(side="left", padx=5)
        # Confirm before stopping the Auto-Read monitor
        def _confirm_stop(d=dispatch, f=frame):
            if messagebox.askyesno("Confirm", "Are you sure you want to close this Auto-Read monitor?"):
                self.stop_monitoring(d, f)
        stop_btn = ctk.CTkButton(frame, text="âŒ", width=30, height=25,
                                 fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                                 command=_confirm_stop)
        stop_btn.pack(side="right", padx=10)

        # â”€â”€â”€ schedule_read_status ensures UI updates run on the Tk main thread â”€â”€â”€
        def schedule_read_status(loc, st, lbl=status_lbl, light=light):
            self.after(0, self.update_status, loc, st, lbl, light)

        mon = CarrierSeleniumMonitor(
            username=self.username, password=self.password,
            dispatch_location=dispatch, stop_event=stop_evt,
            message_types=types,
            reply_mode=self.reply_on_pr_var.get(),
            early_pu_reply=self.reply_on_epu_var.get(),
            auto_walkup=self.auto_walkup_var.get(),
            update_status=schedule_read_status,
            finish_callback=self.remove_monitor,
            log_callback=self.log_action
            , mfa_factor_callback=self.prompt_mfa_factor
            , mfa_code_callback  =self.prompt_mfa_code
            , agent_display_name=self.agent_display_name
        )
        t = threading.Thread(target=mon.run, daemon=True)
        self.running_monitors[dispatch] = {
            "thread": t,
            "stop_event": stop_evt,
            "status_label": status_lbl,
            "light": light,         # â† add this line
            "monitor": mon
        }
        t.start()
        # record an initial "last update" so the Auto-Read watchdog knows it's alive
        self._last_read_update[dispatch] = datetime.now()
        # immediately open the Dropbox tab on this new monitor
        self._ensure_dropbox_tab()
        # Start dropbox refresh cycle if not already running
        self._start_dropbox_refresh_cycle()

    def stop_monitoring(self, loc, frame):
        info = self.running_monitors.pop(loc, None)
        if info:
            # 1) stop its thread & driver
            info["stop_event"].set()
            drv = info["monitor"].driver
            try:
                if drv:
                    drv.quit()
            except:
                pass

            # 2) if that driver held our Dropbox tab, forget it
            if self.dropbox_driver is drv:
                self.dropbox_driver = None
                self.dropbox_handle = None

            # 3) tear down the UI row
            if frame:
                frame.destroy()

        # 4) immediately open a Dropbox tab on the next live monitor
        self._ensure_dropbox_tab()
        # Restart dropbox refresh cycle if we still have monitors
        self._start_dropbox_refresh_cycle()

    def update_status(self, loc, status, label, light):
        # if the UI was torn down, skip updating
        if not (label.winfo_exists() and light.winfo_exists()):
            return
            
        self._last_read_update[loc] = datetime.now()
        st = status.lower()
        if "login failed" in st:
            # show "incorrect password" then go back to login
            self.after(0, lambda: self.handle_login_failure(
                "Incorrect username or password.\nPlease try again."
            ))
            return
        if st == IDLE_STATUS_TEXT.lower():
            light.itemconfigure(light.oval_id, fill="#00FF00")
        elif "cleared" in st or "sent" in st:
            light.itemconfigure(light.oval_id, fill="yellow")
        elif "error" in st or "login" in st:
            light.itemconfigure(light.oval_id, fill="red")
        else:
            light.itemconfigure(light.oval_id, fill=BRAND_ORANGE)
        label.configure(text=f"Status: {status}")

    def remove_monitor(self, loc):
        self.running_monitors.pop(loc, None)

    def _on_clear_reply_choice_change(self, choice: str):
        """
        When the dispatcher selects a canned clear message:
        - For canned options, hide the custom text box and mirror the choice into it.
        - For 'Custom', show the text box so the dispatcher can type their own reply.
        """
        try:
            if choice != "Custom":
                # Keep entry value in sync, but hide the box
                self.clear_reply_entry.delete(0, "end")
                self.clear_reply_entry.insert(0, choice)
                try:
                    self.clear_reply_entry.pack_forget()
                except Exception:
                    pass
            else:
                # Show the box for a custom message if it's not already visible
                # Only pack if it's not currently managed by geometry manager
                if not self.clear_reply_entry.winfo_ismapped():
                    self.clear_reply_entry.pack(side="left", padx=(10, 5))
        except Exception as e:
            logging.warning(f"âš ï¸ Failed to apply clear-reply choice {choice!r}: {e!r}")

    def _on_toggle_clear_name_editor(self):
        """
        Show/hide the CXPC Name editor when the user clicks 'Change'.
        """
        try:
            if self.clear_name_frame is None:
                return
            if self.clear_name_frame.winfo_ismapped():
                self.clear_name_frame.pack_forget()
            else:
                # Show directly under the signature row and above Add Location
                kwargs = {"pady": (0, 8)}
                if hasattr(self, "clear_add_location_button") and self.clear_add_location_button:
                    kwargs["before"] = self.clear_add_location_button
                self.clear_name_frame.pack(**kwargs)
        except Exception as e:
            logging.warning(f"âš ï¸ Failed to toggle CXPC name editor: {e!r}")

    # --- Auto-Clear methods (modified) ---
    def _on_save_clear_name(self):
        """
        Allow dispatcher to manually set their CXPC display name
        (used when EmployeeDirectory lookup fails or to override it).
        """
        try:
            if self.clear_name_entry is None:
                return
            raw = self.clear_name_entry.get().strip()
            if not raw:
                messagebox.showerror(
                    "Missing Name",
                    "Enter your name (for example 'Shawn Williams' or 'Shawn W.') before saving."
                )
                return
            # Reuse the same normalization logic we use for EmployeeDirectory names
            self._set_agent_display_name_from_full(raw)
        except Exception as e:
            logging.warning(f"âš ï¸ Failed to save CXPC name: {e!r}")

    def _start_routr_extension_server(self):
        """Start local HTTP server so RoutR Chrome Extension can check on/off and get signature."""
        if self._routr_extension_server is not None:
            return
        _RoutRExtensionHandler.app_ref = self
        try:
            self._routr_extension_server = _RoutRExtensionServer(
                ('127.0.0.1', ROUTR_EXTENSION_PORT),
                _RoutRExtensionHandler,
            )
        except OSError as e:
            logging.warning(f"âš ï¸ RoutR extension server could not bind to port {ROUTR_EXTENSION_PORT}: {e}")
            return
        def run_server():
            try:
                self._routr_extension_server.serve_forever()
            except Exception:
                pass
        self._routr_extension_thread = threading.Thread(target=run_server, daemon=True)
        self._routr_extension_thread.start()
        logging.info(f"âœ… RoutR Chrome Extension server listening on 127.0.0.1:{ROUTR_EXTENSION_PORT}")

    def _stop_routr_extension_server(self):
        """Stop the local server when ROUTR closes so the extension stops working."""
        if self._routr_extension_server is None:
            return
        try:
            self._routr_extension_server.shutdown()
            self._routr_extension_server = None
        except Exception:
            pass

    def _on_toggle_routr_extension(self):
        """Toggle extension on/off; keep state in routr_extension_enabled."""
        try:
            # Prefer the switch's bound variable if available
            if hasattr(self, "routr_extension_var"):
                self.routr_extension_enabled = bool(self.routr_extension_var.get())
            else:
                self.routr_extension_enabled = not self.routr_extension_enabled
        except Exception:
            # Fallback: simple toggle
            self.routr_extension_enabled = not self.routr_extension_enabled

    def _get_routr_extension_dir(self):
        """Path where the RoutR_Chrome_Extension folder should live (user Downloads)."""
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        downloads_dir = os.path.join(user_profile, "Downloads")
        try:
            os.makedirs(downloads_dir, exist_ok=True)
        except Exception:
            pass
        return os.path.join(downloads_dir, "RoutR_Chrome_Extension")

    def _ensure_routr_extension_folder(self):
        """Create/populate RoutR_Chrome_Extension under the current user's Downloads folder."""
        import shutil

        target_dir = self._get_routr_extension_dir()
        os.makedirs(target_dir, exist_ok=True)

        # Determine source folder containing extension files
        if getattr(sys, "frozen", False):
            src_base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        else:
            src_base = os.path.dirname(os.path.abspath(__file__))

        src_dir = os.path.join(src_base, "RoutR_Chrome_Extension")
        if not os.path.isdir(src_dir):
            messagebox.showerror(
                "Extension files missing",
                "Could not find the RoutR_Chrome_Extension template folder.\n\n"
                "Expected at:\n"
                f"{src_dir}\n\n"
                "Make sure the folder exists in development and is bundled with the EXE."
            )
            return False

        try:
            for name in os.listdir(src_dir):
                src_path = os.path.join(src_dir, name)
                dst_path = os.path.join(target_dir, name)
                if os.path.isdir(src_path):
                    shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
                else:
                    shutil.copy2(src_path, dst_path)

            # Also drop a copy of the ROUTR logo into the extension folder
            # so Chrome can use it as the extension icon (overwrites placeholder).
            logo_src = None
            for candidate in ("ROUTR LOGO.png", "ROUTR_LOGO.png"):
                p = os.path.join(src_base, candidate)
                if os.path.isfile(p):
                    logo_src = p
                    break
            if logo_src:
                logo_dst = os.path.join(target_dir, "routr_icon_128.png")
                try:
                    shutil.copy2(logo_src, logo_dst)
                except Exception as e:
                    logging.warning("Could not overwrite extension icon with logo: %s", e)
            # Template must include routr_icon_128.png so the extension loads for all users
            if not os.path.isfile(os.path.join(target_dir, "routr_icon_128.png")):
                messagebox.showerror(
                    "Extension incomplete",
                    "The extension folder is missing routr_icon_128.png.\n\n"
                    "Include routr_icon_128.png in the RoutR_Chrome_Extension template folder\n"
                    "and in your EXE bundle so Chrome can load the extension."
                )
                return False

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to copy extension files:\n{e}"
            )
            return False

        messagebox.showinfo(
            "RoutR Chrome Extension",
            f"Extension folder is ready at:\n{target_dir}"
        )
        return True

    def _open_chrome_extensions_page(self):
        """Try to open Chrome directly to the extensions page so the user can click Load unpacked."""
        import subprocess
        paths = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
        ]
        for exe in paths:
            if exe and os.path.isfile(exe):
                try:
                    # Open a fresh window directly to the extensions page
                    subprocess.Popen([exe, "--new-window", "chrome://extensions"], shell=False)
                    return True
                except Exception:
                    pass
        return False

    def _on_install_routr_extension(self):
        """Open extension folder, open Chrome to extensions page, and show short instructions."""
        ext_dir = self._get_routr_extension_dir()
        if not os.path.isdir(ext_dir):
            # Try to create/populate the folder first
            if not self._ensure_routr_extension_folder():
                return
            ext_dir = self._get_routr_extension_dir()
            if not os.path.isdir(ext_dir):
                messagebox.showerror(
                    "Extension folder not found",
                    f"RoutR Chrome Extension folder not found at:\n{ext_dir}\n\n"
                    "Ensure the 'RoutR_Chrome_Extension' folder is bundled next to ROUTR."
                )
                return
        try:
            os.startfile(ext_dir)
        except Exception:
            try:
                import webbrowser
                webbrowser.open("file:///" + ext_dir.replace("\\", "/"))
            except Exception:
                messagebox.showerror("Error", "Could not open the extension folder.")
                return

        # Do not try to open chrome://extensions automatically; many managed
        # environments override startup URLs. Just give clear manual steps.
        messagebox.showinfo(
            "Install RoutR Chrome Extension",
            "The extension folder is open.\n\n"
            "In Chrome:\n"
            "1. Open a new tab and go to chrome://extensions\n"
            "2. Turn on 'Developer mode' (top right) if it's off\n"
            "3. Click 'Load unpacked'\n"
            "4. Select the folder that just opened (RoutR_Chrome_Extension)\n\n"
            "Use the 'Extension' switch in Utility to enable/disable the signature when ROUTR is open."
        )

    def add_clear_location(self, with_console=False):
        if with_console and not getattr(self, "console_allocated_clear", False):
            import ctypes, sys
            ctypes.windll.kernel32.AllocConsole()
            sys.stdout = open("CONOUT$", "w", buffering=1)
            sys.stderr = open("CONOUT$", "w", buffering=1)
            h = logging.StreamHandler(sys.stdout)
            h.setFormatter(logging.Formatter("%(message)s"))
            logging.getLogger().addHandler(h)
            self.console_allocated_clear = True
            logging.info("ðŸ” Console allocated for Auto-Clear")

        loc = self.clear_dispatch_entry.get().strip().upper()
        reply = self.clear_reply_entry.get().strip()

        # Require a CXPC display name before starting Auto-Clear
        if not self.agent_display_name:
            messagebox.showerror(
                "Missing CXPC Name",
                "Please set your CXPC name before starting Auto-Clear.\n\n"
                "Either:\n"
                " â€¢ Log in so EmployeeDirectory can detect your name, or\n"
                " â€¢ Type your name in the 'CXPC Name' box on the Auto-Clear tab and click 'Save Name'."
            )
            return
        # initialize skip list for this location
        self.skip_clear_routes.setdefault(loc, set())
        if not loc or not reply:
            messagebox.showerror("Missing Info", "Enter dispatch location and clear reply.")
            return
        if len(self.running_clear_monitors) >= 5:
            messagebox.showerror("Limit Reached", "You may monitor up to 5 clear-locations.")
            return
        if loc in self.running_clear_monitors:
            messagebox.showerror("Duplicate", "Auto-Clear monitor already running.")
            return
        # â”€â”€â”€ Use green-light check instead of text match â”€â”€â”€
        for m in self.running_clear_monitors.values():
            light = m["light"]
            fill = light.itemcget(light.oval_id, "fill")
            if fill != "#00FF00":  # not green
                messagebox.showerror("Busy", "Wait until clear monitors are idle.")
                return

        self.clear_dispatch_entry.delete(0, ctk.END)
        frame = ctk.CTkFrame(self.clear_locations_container, fg_color="black")
        frame.pack(fill="x", pady=5)

        ctk.CTkLabel(frame, text=f"Location: {loc}",
                     font=ctk.CTkFont(size=12), text_color="white")\
            .pack(side="left", padx=10)

        # Abbreviated reply label + tooltip
        abbr = reply if len(reply) <= 5 else (reply[:5] + "â€¦")
        reply_lbl = ctk.CTkLabel(frame, text=f"Reply: {abbr}",
                                 font=ctk.CTkFont(size=12), text_color="white")
        reply_lbl.pack(side="left", padx=10)
        if abbr != reply:
            ToolTip(reply_lbl, reply)

        status_lbl = ctk.CTkLabel(frame, text="Status: Startingâ€¦",
                                  font=ctk.CTkFont(size=12), text_color="white")
        status_lbl.pack(side="left", padx=10)

        light = ctk.CTkCanvas(frame, width=12, height=12, bg="black", highlightthickness=0)
        light.oval_id = light.create_oval(
            0, 0, 12, 12,
            fill=BRAND_ORANGE, outline=BRAND_ORANGE
        )    
        light.pack(side="left", padx=5)

        # Confirm before stopping the Auto-Clear monitor
        def _confirm_clear_stop(l=loc, f=frame):
            if messagebox.askyesno("Confirm", "Are you sure you want to close this Auto-Clear monitor?"):
                self.stop_clear_monitor(l, f)
        stop_btn = ctk.CTkButton(frame, text="âŒ", width=30, height=25,
                                 fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                                 command=_confirm_clear_stop)
        stop_btn.pack(side="right", padx=10)

        # â”€â”€â”€ Audit button (failed clears) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        audit_btn = ctk.CTkButton(
            frame, text="Audit", width=60, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=lambda l=loc: self.show_audit_window(l)
        )
        audit_btn.pack(side="right", padx=(5, 0))
        # â”€â”€â”€ Not-Clear skip toggle â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        nc_btn = ctk.CTkButton(
            frame, text="S/C", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=lambda l=loc: self.show_skip_popup(l)
        )
        nc_btn.pack(side="right", padx=(5, 0))
        ToolTip(nc_btn, "Skip Clear")

        # â”€â”€â”€ History button â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        history_btn = ctk.CTkButton(
            frame, text="History", width=60, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=lambda l=loc: self.show_clear_history(l)
        )
        # Place History closer to the center so it doesn't hide under the scrollbar
        history_btn.pack(side="left", padx=(10, 0))

        stop_evt = threading.Event()
        # â”€â”€â”€ schedule_clear_status ensures UI updates run on the Tk main thread â”€â”€â”€
        def schedule_clear_status(l, s, lbl=status_lbl, light=light):
            self.after(0, self._update_clear_status, l, s, lbl, light)

        mon = CarrierClearSeleniumMonitor(
            username=self.username, password=self.password,
            dispatch_location=loc, clear_reply=reply,
            stop_event=stop_evt,
            # â† add this argument so the monitor sees your live set
            skip_routes=self.skip_clear_routes[loc],
            update_status=schedule_clear_status,
            finish_callback=lambda l: self.stop_clear_monitor(l, None),
            log_callback=self.log_action,
            mfa_factor_callback=self.prompt_mfa_factor,
            mfa_code_callback  =self.prompt_mfa_code,
            agent_display_name=self.agent_display_name
        )
        t = threading.Thread(target=mon.run, daemon=True)
        self.running_clear_monitors[loc] = {
            "thread": t, "stop_event": stop_evt,
            "status_label": status_lbl, "light": light,
            "frame": frame, "monitor": mon
        }
        t.start()
        # record an initial "last update" so watchdog will notice if run() hangs
        self._last_clear_update[loc] = datetime.now()

    def stop_clear_monitor(self, loc, frame):
        info = self.running_clear_monitors.pop(loc, None)
        if info:
            info["stop_event"].set()
            try:
                drv = info["monitor"].driver
                if drv:
                    drv.quit()
            except:
                pass
            if frame:
                frame.destroy()
        # also stop & clean up the audit monitor for this location
        self.stop_audit_monitor(loc)
        # remove any pending audit routes
        self.clear_audit_routes.pop(loc, None)

    def _update_clear_status(self, loc, status, lbl, light):
        # skip if widgets no longer exist
        if not (lbl.winfo_exists() and light.winfo_exists()):
            return

        self._last_clear_update[loc] = datetime.now()
        if "login failed" in status.lower():
            self.after(0, lambda: self.handle_login_failure(
                "Incorrect username or password.\nPlease try again."
            ))
            return
        # truncate to keep the row compact so buttons stay visible
        # use a very short prefix and aggressive truncation
        max_len = 10
        abbr = status if len(status) <= max_len else status[:max_len] + "â€¦"
        lbl.configure(text=f"St: {abbr}")

        # attach tooltip if truncated
        if len(status) > max_len:
            ToolTip(lbl, status)

        # existing light-color logic
        st = status.lower()
        if st == CLEAR_IDLE_STATUS_TEXT.lower():
            light.itemconfigure(light.oval_id, fill="#00FF00")
        elif "cleared" in st or "sent" in st:
            light.itemconfigure(light.oval_id, fill="yellow")
        else:
            light.itemconfigure(light.oval_id, fill=BRAND_ORANGE)

        # â”€â”€â”€ capture any "open stops" events for auditing â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if "open stops" in st:
            mon = self.running_clear_monitors.get(loc, {}).get("monitor")
            if mon:
                # snapshot and then clear the monitor's reported-routes buffer
                new_routes = set(mon.open_routes_reported)
                logging.info(f"ðŸ” Audit check for {loc}, routes: {new_routes}")
                mon.open_routes_reported.clear()
                # only add the truly new ones
                for route in new_routes:
                    self.clear_audit_routes.setdefault(loc, set()).add(route)
                # if we actually got any new routes, start the audit monitor
                if new_routes:
                    # use the app's log_action, not a non-existent log_callback
                    logging.info(f"âœ… Adding {len(new_routes)} route(s) to audit: {sorted(new_routes)}")
                    self.log_action(f"{loc}: routes in audit: {sorted(new_routes)}")
                    self.start_audit_monitor(loc)
                else:
                    logging.debug(f"â„¹ï¸ No new routes found in open_routes_reported for {loc}")
            else:
                logging.warning(f"âš ï¸ No monitor found for {loc} when checking for audit routes")

    def show_clear_history(self, loc):
        info = self.running_clear_monitors.get(loc)
        if not info:
            messagebox.showinfo("History", f"No monitor for {loc}.")
            return
        hist = info["monitor"].history
        if not hist:
            messagebox.showinfo("History", f"No routes cleared for {loc}.")
            return

        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.grab_set()
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(header, text=f"Cleared Routes for {loc}",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=BRAND_ORANGE).pack(side="left", padx=10, pady=5)
        def close_popup():
            popup.destroy()
        ctk.CTkButton(header, text="âœ•", width=30, height=25,
                      fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                      command=close_popup).pack(side="right", padx=10, pady=5)
        def start_move(event):
            popup.x0 = event.x; popup.y0 = event.y
        def do_move(event):
            x = popup.winfo_x() + (event.x - popup.x0)
            y = popup.winfo_y() + (event.y - popup.y0)
            popup.geometry(f"+{x}+{y}")
        header.bind("<ButtonPress-1>", start_move)
        header.bind("<B1-Motion>", do_move)

        # content area with ZERO vertical padding
        content = ctk.CTkScrollableFrame(inner, fg_color="black", height=200)
        content.pack(fill="both", padx=10, pady=(0,0))

        # pack each route label directly, with zero top/bottom padding
        for r, ts in hist:
            ctk.CTkLabel(
                content,
                text=f"Route {r} at {ts:%Y-%m-%d %H:%M:%S}",
                font=ctk.CTkFont(size=12),
                text_color="white"
            ).pack(fill="x", anchor="w", padx=10, pady=(0,0))

        self.update_idletasks()
        mx, my = self.winfo_x(), self.winfo_y()
        mw, mh = self.winfo_width(), self.winfo_height()
        popup.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        x = mx + (mw - pw)//2; y = my + (mh - ph)//2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")

    def show_una_popup(self):
        """Show popup with UNA dispatch information."""
        try:
            # Check if we're in widget context
            if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
                parent_window = self.widget_window
            else:
                parent_window = self
            
            if not hasattr(self, 'una_popup_data') or not self.una_popup_data:
                messagebox.showinfo("UNA Dispatches", "No UNA dispatches found.")
                return

            popup = ctk.CTkToplevel(parent_window)
            popup.overrideredirect(True)
            popup.grab_set()
            # Make popup appear above all windows including ROUTR
            popup.attributes("-topmost", True)
            outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
            outer.pack(fill="both", expand=True)
            inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
            inner.pack(fill="both", expand=True, padx=2, pady=2)

            header = ctk.CTkFrame(inner, fg_color="black")
            header.pack(fill="x")
            ctk.CTkLabel(header, text="UNA Dispatches",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color="#FF6600").pack(side="left", padx=10, pady=5)
            def close_popup():
                popup.destroy()
            def copy_to_clipboard():
                if self.una_popup_data:
                    copy_text = f"UNA Dispatches ({len(self.una_popup_data)}):\n" + "\n".join(self.una_popup_data)
                else:
                    copy_text = "No UNA dispatches found"
                popup.clipboard_clear()
                popup.clipboard_append(copy_text)
                popup.update()  # Ensure clipboard is updated
            
            ctk.CTkButton(header, text="âœ•", width=30, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=close_popup).pack(side="right", padx=(5,10), pady=5)
            ctk.CTkButton(header, text="Copy", width=50, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=copy_to_clipboard).pack(side="right", padx=0, pady=5)
            def start_move(event):
                popup.x0 = event.x; popup.y0 = event.y
            def do_move(event):
                x = popup.winfo_x() + (event.x - popup.x0)
                y = popup.winfo_y() + (event.y - popup.y0)
                popup.geometry(f"+{x}+{y}")
            header.bind("<ButtonPress-1>", start_move)
            header.bind("<B1-Motion>", do_move)

            content = ctk.CTkScrollableFrame(inner, fg_color="black", height=300, width=400)
            content.pack(fill="both", padx=10, pady=(0,0))

            for disp in self.una_popup_data[:100]:  # Limit to first 100
                ctk.CTkLabel(
                    content,
                    text=disp,
                    font=ctk.CTkFont(size=12),
                    text_color="white"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))

            if len(self.una_popup_data) > 100:
                ctk.CTkLabel(
                    content,
                    text=f"... and {len(self.una_popup_data) - 100} more",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))

            self._position_popup(popup)
        except Exception as e:
            logging.error(f"Error showing UNA popup: {e}")
            import traceback
            traceback.print_exc()
            try:
                messagebox.showerror("Error", f"Failed to show UNA popup: {e}")
            except:
                pass

    def _enable_fast_scroll(self, scrollable, lines=6):
        """Increase mouse-wheel scroll speed for a CTkScrollableFrame and ensure it works over empty space.
        Uses enter/leave to bind wheel events to the toplevel while the cursor is inside the scrollable area.
        """
        try:
            canvas = getattr(scrollable, "_parent_canvas", None)
            if not canvas:
                for child in scrollable.winfo_children():
                    if isinstance(child, tk.Canvas):
                        canvas = child
                        break
            if not canvas:
                return
            def on_wheel(event, c=canvas, step_lines=lines):
                try:
                    direction = -1 if event.delta > 0 else 1
                except Exception:
                    direction = 1
                c.yview_scroll(direction * step_lines, "units")
                return "break"

            # Bind wheel events while pointer is inside the scrollable region
            def _bind_all(_=None):
                try:
                    scrollable.bind_all("<MouseWheel>", on_wheel, add=True)
                    scrollable.bind_all("<Button-4>", lambda e: on_wheel(type("E", (), {"delta": 120})()))
                    scrollable.bind_all("<Button-5>", lambda e: on_wheel(type("E", (), {"delta": -120})()))
                except Exception:
                    pass
            def _unbind_all(_=None):
                try:
                    scrollable.unbind_all("<MouseWheel>")
                    scrollable.unbind_all("<Button-4>")
                    scrollable.unbind_all("<Button-5>")
                except Exception:
                    pass

            scrollable.bind("<Enter>", _bind_all)
            scrollable.bind("<Leave>", _unbind_all)
            canvas.bind("<Enter>", _bind_all)
            canvas.bind("<Leave>", _unbind_all)

            # Also bind the internal content frame if present to capture gaps
            inner = getattr(scrollable, "_scrollable_frame", None)
            if inner:
                inner.bind("<Enter>", _bind_all)
                inner.bind("<Leave>", _unbind_all)
        except Exception:
            pass

    def show_early_popup(self):
        """Show popup with early pickup information."""
        try:
            # Check if we're in widget context
            if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
                parent_window = self.widget_window
            else:
                parent_window = self
            
            if not hasattr(self, 'early_popup_data') or not self.early_popup_data:
                messagebox.showinfo("Early Pickups", "No early pickups found.")
                return

            popup = ctk.CTkToplevel(parent_window)
            popup.overrideredirect(True)
            popup.grab_set()
            # Make popup appear above all windows including ROUTR
            popup.attributes("-topmost", True)
            outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
            outer.pack(fill="both", expand=True)
            inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
            inner.pack(fill="both", expand=True, padx=2, pady=2)

            header = ctk.CTkFrame(inner, fg_color="black")
            header.pack(fill="x")
            ctk.CTkLabel(header, text="Early Pickups",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color="#FF6600").pack(side="left", padx=10, pady=5)
            def close_popup():
                popup.destroy()
            def copy_to_clipboard():
                if self.early_popup_data:
                    lines = []
                    route_groups = {}
                    for pickup in self.early_popup_data:
                        if " -> " in pickup:
                            disp_num, route = pickup.split(" -> ", 1)
                            if route not in route_groups:
                                route_groups[route] = []
                            route_groups[route].append(disp_num)
                        else:
                            if "Unknown" not in route_groups:
                                route_groups["Unknown"] = []
                            route_groups["Unknown"].append(pickup)
                    
                    for route, disp_list in route_groups.items():
                        lines.append(f"Route {route}: {', '.join(disp_list)}")
                    copy_text = f"Early Pickups ({len(self.early_popup_data)}):\n" + "\n".join(lines)
                else:
                    copy_text = "No early pickups found"
                popup.clipboard_clear()
                popup.clipboard_append(copy_text)
                popup.update()
            
            ctk.CTkButton(header, text="âœ•", width=30, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=close_popup).pack(side="right", padx=(5,10), pady=5)
            ctk.CTkButton(header, text="Copy", width=50, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=copy_to_clipboard).pack(side="right", padx=0, pady=5)
            def start_move(event):
                popup.x0 = event.x; popup.y0 = event.y
            def do_move(event):
                x = popup.winfo_x() + (event.x - popup.x0)
                y = popup.winfo_y() + (event.y - popup.y0)
                popup.geometry(f"+{x}+{y}")
            header.bind("<ButtonPress-1>", start_move)
            header.bind("<B1-Motion>", do_move)

            content = ctk.CTkScrollableFrame(inner, fg_color="black", height=300, width=400)
            content.pack(fill="both", padx=10, pady=(0,0))

            # Group early pickups by route
            route_groups = {}
            for pickup in self.early_popup_data[:100]:  # Limit to first 100
                # Extract route and disp# from pickup string (format: "Disp# -> Route")
                if " -> " in pickup:
                    disp_num, route = pickup.split(" -> ", 1)
                    if route not in route_groups:
                        route_groups[route] = []
                    route_groups[route].append(disp_num)
                else:
                    # Fallback for unexpected format
                    if "Unknown" not in route_groups:
                        route_groups["Unknown"] = []
                    route_groups["Unknown"].append(pickup)

            for route, disp_list in route_groups.items():
                # Display route header with dispatch numbers on same line
                ctk.CTkLabel(
                    content,
                    text=f"Route {route}: {', '.join(disp_list)}",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(2,0))

            if len(self.early_popup_data) > 100:
                ctk.CTkLabel(
                    content,
                    text=f"... and {len(self.early_popup_data) - 100} more",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))

            self._position_popup(popup)
        except Exception as e:
            logging.error(f"Error showing early popup: {e}")
            import traceback
            traceback.print_exc()
            try:
                messagebox.showerror("Error", f"Failed to show early popup: {e}")
            except:
                pass

    def show_late_popup(self):
        """Show popup with late pickup information."""
        try:
            # Check if we're in widget context
            if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
                parent_window = self.widget_window
            else:
                parent_window = self
            
            if not hasattr(self, 'late_popup_data') or not self.late_popup_data:
                messagebox.showinfo("Late Pickups", "No late pickups found.")
                return

            popup = ctk.CTkToplevel(parent_window)
            popup.overrideredirect(True)
            popup.grab_set()
            # Make popup appear above all windows including ROUTR
            popup.attributes("-topmost", True)
            outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
            outer.pack(fill="both", expand=True)
            inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
            inner.pack(fill="both", expand=True, padx=2, pady=2)

            header = ctk.CTkFrame(inner, fg_color="black")
            header.pack(fill="x")
            ctk.CTkLabel(header, text="Late Pickups",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color="#FF6600").pack(side="left", padx=10, pady=5)
            def close_popup():
                popup.destroy()
            def copy_to_clipboard():
                if self.late_popup_data:
                    lines = []
                    route_groups = {}
                    for pickup in self.late_popup_data:
                        if " -> " in pickup:
                            disp_num, route = pickup.split(" -> ", 1)
                            if route not in route_groups:
                                route_groups[route] = []
                            route_groups[route].append(disp_num)
                        else:
                            if "Unknown" not in route_groups:
                                route_groups["Unknown"] = []
                            route_groups["Unknown"].append(pickup)
                    
                    for route, disp_list in route_groups.items():
                        lines.append(f"Route {route}: {', '.join(disp_list)}")
                    copy_text = f"Late Pickups ({len(self.late_popup_data)}):\n" + "\n".join(lines)
                else:
                    copy_text = "No late pickups found"
                popup.clipboard_clear()
                popup.clipboard_append(copy_text)
                popup.update()
            
            ctk.CTkButton(header, text="âœ•", width=30, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=close_popup).pack(side="right", padx=(5,10), pady=5)
            ctk.CTkButton(header, text="Copy", width=50, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=copy_to_clipboard).pack(side="right", padx=0, pady=5)
            def start_move(event):
                popup.x0 = event.x; popup.y0 = event.y
            def do_move(event):
                x = popup.winfo_x() + (event.x - popup.x0)
                y = popup.winfo_y() + (event.y - popup.y0)
                popup.geometry(f"+{x}+{y}")
            header.bind("<ButtonPress-1>", start_move)
            header.bind("<B1-Motion>", do_move)

            content = ctk.CTkScrollableFrame(inner, fg_color="black", height=300, width=400)
            content.pack(fill="both", padx=10, pady=(0,0))

            # Group late pickups by route
            route_groups = {}
            for pickup in self.late_popup_data[:100]:  # Limit to first 100
                # Extract route and disp# from pickup string (format: "Disp# -> Route")
                if " -> " in pickup:
                    disp_num, route = pickup.split(" -> ", 1)
                    if route not in route_groups:
                        route_groups[route] = []
                    route_groups[route].append(disp_num)
                else:
                    # Fallback for unexpected format
                    if "Unknown" not in route_groups:
                        route_groups["Unknown"] = []
                    route_groups["Unknown"].append(pickup)

            for route, disp_list in route_groups.items():
                # Display route header with dispatch numbers on same line
                ctk.CTkLabel(
                    content,
                    text=f"Route {route}: {', '.join(disp_list)}",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(2,0))

            if len(self.late_popup_data) > 100:
                ctk.CTkLabel(
                    content,
                    text=f"... and {len(self.late_popup_data) - 100} more",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))

            self._position_popup(popup)
        except Exception as e:
            logging.error(f"Error showing late popup: {e}")
            import traceback
            traceback.print_exc()
            try:
                messagebox.showerror("Error", f"Failed to show late popup: {e}")
            except:
                pass

    def show_untransmitted_popup(self):
        """Show popup with untransmitted routes information."""
        try:
            # Check if we're in widget context
            if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
                parent_window = self.widget_window
            else:
                parent_window = self
            
            if not hasattr(self, 'untransmitted_popup_data') or not self.untransmitted_popup_data:
                messagebox.showinfo("Untransmitted Stops", "No untransmitted stops found.")
                return

            popup = ctk.CTkToplevel(parent_window)
            popup.overrideredirect(True)
            popup.grab_set()
            # Make popup appear above all windows including ROUTR
            popup.attributes("-topmost", True)
            outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
            outer.pack(fill="both", expand=True)
            inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
            inner.pack(fill="both", expand=True, padx=2, pady=2)

            header = ctk.CTkFrame(inner, fg_color="black")
            header.pack(fill="x")
            ctk.CTkLabel(header, text="Untransmitted Stops",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color="#FF6600").pack(side="left", padx=10, pady=5)
            def close_popup():
                popup.destroy()
            def copy_to_clipboard():
                if hasattr(self, 'organized_untransmitted_data'):
                    lines = []
                    lines.append("=== UNTRANSMITTED ROUTES ===")
                    
                    # Add unsafe routes
                    for route, count in self.organized_untransmitted_data['unsafe'].items():
                        lines.append(f"âš  Route {route} - {count}")
                    
                    # Add safe routes
                    for route, count in self.organized_untransmitted_data['safe'].items():
                        lines.append(f"âœ“ Route {route} - {count}")
                    
                    copy_text = "\n".join(lines)
                else:
                    copy_text = "No untransmitted stops found"
                popup.clipboard_clear()
                popup.clipboard_append(copy_text)
                popup.update()
            
            ctk.CTkButton(header, text="âœ•", width=30, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=close_popup).pack(side="right", padx=(5,10), pady=5)
            ctk.CTkButton(header, text="Copy", width=50, height=25,
                          fg_color="#FF6600", hover_color="#800080",
                          command=copy_to_clipboard).pack(side="right", padx=0, pady=5)
            
            def start_move(event):
                popup.x0 = event.x; popup.y0 = event.y
            def do_move(event):
                x = popup.winfo_x() + (event.x - popup.x0)
                y = popup.winfo_y() + (event.y - popup.y0)
                popup.geometry(f"+{x}+{y}")
            header.bind("<ButtonPress-1>", start_move)
            header.bind("<B1-Motion>", do_move)

            # Create single content area
            content = ctk.CTkScrollableFrame(inner, fg_color="black", height=300, width=400)
            content.pack(fill="both", expand=True, padx=10, pady=(0,10))
            
            # Initialize organized untransmitted data if not exists
            if not hasattr(self, 'organized_untransmitted_data'):
                self.organized_untransmitted_data = {
                    'unsafe': self.untransmitted_popup_data.copy(),
                    'safe': {}
                }
            else:
                # Preserve existing organization, only add NEW routes that don't exist in either tab
                new_routes = {}
                for route, count in self.untransmitted_popup_data.items():
                    if (route not in self.organized_untransmitted_data['unsafe'] and 
                        route not in self.organized_untransmitted_data['safe']):
                        new_routes[route] = count
                
                # Add new routes to unsafe tab
                for route, count in new_routes.items():
                    self.organized_untransmitted_data['unsafe'][route] = count
                
                # Remove routes that are no longer active (no longer in source data)
                all_source_routes = set(self.untransmitted_popup_data.keys())
                
                # Clean up unsafe tab
                for route in list(self.organized_untransmitted_data['unsafe'].keys()):
                    if route not in all_source_routes:
                        del self.organized_untransmitted_data['unsafe'][route]
                
                # Clean up safe tab
                for route in list(self.organized_untransmitted_data['safe'].keys()):
                    if route not in all_source_routes:
                        del self.organized_untransmitted_data['safe'][route]

            # --- Sync counts from the latest source data for existing routes ---
            try:
                for route, count in self.untransmitted_popup_data.items():
                    if route in self.organized_untransmitted_data['unsafe']:
                        self.organized_untransmitted_data['unsafe'][route] = count
                    if route in self.organized_untransmitted_data['safe']:
                        self.organized_untransmitted_data['safe'][route] = count
            except Exception:
                pass
            

            
            # Populate all routes in single view
            all_routes = []
            
            # Add unsafe routes
            for route, count in self.organized_untransmitted_data['unsafe'].items():
                all_routes.append((route, count, "unsafe"))
            
            # Add safe routes
            for route, count in self.organized_untransmitted_data['safe'].items():
                all_routes.append((route, count, "safe"))
            
            # Sort all routes
            sorted_routes = sorted(all_routes, key=lambda x: self.safe_route_sort_key(x[0]))
            
            for route, count, status in sorted_routes[:200]:  # Limit to first 200
                # Create frame for route
                route_frame = ctk.CTkFrame(content, fg_color="black")
                route_frame.pack(fill="x", padx=10, pady=(0,2))
                
                # Route label with status indicator
                status_color = "#00AA00" if status == "safe" else "white"
                status_text = "âœ“" if status == "safe" else "âš "
                
                ctk.CTkLabel(
                    route_frame,
                    text=f"{status_text} Route {route} - {count}",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=status_color
                ).pack(side="left", fill="x", expand=True, anchor="w", padx=(10,0))

            if len(sorted_routes) > 200:
                ctk.CTkLabel(
                    content,
                    text=f"... and {len(sorted_routes) - 200} more routes",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))
            
            # Store references for refresh functionality
            self.untransmitted_popup_content = content
            
            # Store popup reference for refresh functionality
            self.untransmitted_popup = popup
            
            # Store the current untransmitted data for future refreshes
            if not hasattr(self, 'untransmitted_popup_data'):
                self.untransmitted_popup_data = {}
            
            # Position popup above main app
            self._position_popup(popup)
        except Exception as e:
            logging.error(f"Error showing untransmitted popup: {e}")
            import traceback
            traceback.print_exc()
            try:
                messagebox.showerror("Error", f"Failed to show untransmitted popup: {e}")
            except:
                pass

    def refresh_untransmitted_popup_content(self):
        """Refresh the untransmitted popup content in real-time without closing"""
        try:
            if not hasattr(self, 'untransmitted_popup') or not self.untransmitted_popup:
                return
            
            # IMPORTANT: Update organized data with current source data while preserving manual organization
            # This ensures your manual moves are preserved when new data comes in
            if hasattr(self, 'untransmitted_popup_data'):
                # Only add NEW routes that don't exist in either tab
                new_routes = {}
                for route, count in self.untransmitted_popup_data.items():
                    if (route not in self.organized_untransmitted_data['unsafe'] and 
                        route not in self.organized_untransmitted_data['safe']):
                        new_routes[route] = count
                
                # Add new routes to unsafe tab
                for route, count in new_routes.items():
                    self.organized_untransmitted_data['unsafe'][route] = count
                
                # Remove routes that are no longer active (no longer in source data)
                all_source_routes = set(self.untransmitted_popup_data.keys())
                
                # Clean up unsafe tab
                for route in list(self.organized_untransmitted_data['unsafe'].keys()):
                    if route not in all_source_routes:
                        del self.organized_untransmitted_data['unsafe'][route]
                
                # Clean up safe tab
                for route in list(self.organized_untransmitted_data['safe'].keys()):
                    if route not in all_source_routes:
                        del self.organized_untransmitted_data['safe'][route]

                # --- Sync counts from the latest source data for existing routes ---
                for route, count in self.untransmitted_popup_data.items():
                    if route in self.organized_untransmitted_data['unsafe']:
                        self.organized_untransmitted_data['unsafe'][route] = count
                    if route in self.organized_untransmitted_data['safe']:
                        self.organized_untransmitted_data['safe'][route] = count
            
            # Update header counts
            total_unsafe = len(self.organized_untransmitted_data['unsafe'])
            total_safe = len(self.organized_untransmitted_data['safe'])
            
            # Find and update the header label
            for widget in self.untransmitted_popup.winfo_children():
                if hasattr(widget, 'winfo_children'):
                    for child in widget.winfo_children():
                        if hasattr(child, 'winfo_children'):
                            for grandchild in child.winfo_children():
                                if isinstance(grandchild, ctk.CTkLabel) and "Untransmitted Stops" in grandchild.cget("text"):
                                    grandchild.configure(text=f"Untransmitted Stops (Unsafe: {total_unsafe}, Safe: {total_safe})")
                                    break
            
            # Clear and repopulate single content area
            for widget in self.untransmitted_popup_content.winfo_children():
                widget.destroy()
            
            # Populate all routes in single view
            all_routes = []
            
            # Add unsafe routes
            for route, count in self.organized_untransmitted_data['unsafe'].items():
                all_routes.append((route, count, "unsafe"))
            
            # Add safe routes
            for route, count in self.organized_untransmitted_data['safe'].items():
                all_routes.append((route, count, "safe"))
            
            # Sort all routes
            sorted_routes = sorted(all_routes, key=lambda x: self.safe_route_sort_key(x[0]))
            
            for route, count, status in sorted_routes[:200]:  # Limit to first 200
                # Create frame for route
                route_frame = ctk.CTkFrame(self.untransmitted_popup_content, fg_color="black")
                route_frame.pack(fill="x", padx=10, pady=(0,2))
                
                # Route label with status indicator
                status_color = "#00AA00" if status == "safe" else "white"
                status_text = "âœ“" if status == "safe" else "âš "
                
                ctk.CTkLabel(
                    route_frame,
                    text=f"{status_text} Route {route} - {count}",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=status_color
                ).pack(side="left", fill="x", expand=True, anchor="w", padx=(10,0))

            if len(sorted_routes) > 200:
                ctk.CTkLabel(
                    self.untransmitted_popup_content,
                    text=f"... and {len(sorted_routes) - 200} more routes",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))
            
            # Verify data integrity after refresh
            self.verify_untransmitted_data_integrity()
            
            # Debug logging to help track organization
            print(f"ðŸ” Untransmitted refresh complete - Unsafe: {len(self.organized_untransmitted_data['unsafe'])}, Safe: {len(self.organized_untransmitted_data['safe'])}")
            print(f"ðŸ” Safe routes: {list(self.organized_untransmitted_data['safe'].keys())}")
                                    
        except Exception as e:
            print(f"ðŸ” Error refreshing untransmitted popup: {e}")
    
    def move_to_safe_untransmitted(self, route, count):
        """Move a route from unsafe to safe tab."""
        # Ensure route is completely removed from unsafe
        if route in self.organized_untransmitted_data['unsafe']:
            del self.organized_untransmitted_data['unsafe'][route]
        
        # Ensure route is completely removed from safe before adding (in case it somehow exists in both)
        if route in self.organized_untransmitted_data['safe']:
            del self.organized_untransmitted_data['safe'][route]
        
        # Add to safe tab
        self.organized_untransmitted_data['safe'][route] = count
        
        # Verify data integrity - ensure route only exists in safe
        assert route not in self.organized_untransmitted_data['unsafe'], f"Route {route} still exists in unsafe after move"
        
        self.refresh_untransmitted_popup_content()
    
    def move_to_unsafe_untransmitted(self, route, count):
        """Move a route from safe to unsafe tab."""
        # Ensure route is completely removed from safe
        if route in self.organized_untransmitted_data['safe']:
            del self.organized_untransmitted_data['safe'][route]
        
        # Ensure route is completely removed from unsafe before adding (in case it somehow exists in both)
        if route in self.organized_untransmitted_data['unsafe']:
            del self.organized_untransmitted_data['unsafe'][route]
        
        # Add to unsafe tab
        self.organized_untransmitted_data['unsafe'][route] = count
        
        # Verify data integrity - ensure route only exists in unsafe
        assert route not in self.organized_untransmitted_data['safe'], f"Route {route} still exists in safe after move"
        
        self.refresh_untransmitted_popup_content()
    
    def verify_untransmitted_data_integrity(self):
        """Verify that no routes exist in both safe and unsafe tabs simultaneously."""
        unsafe_routes = set(self.organized_untransmitted_data['unsafe'].keys())
        safe_routes = set(self.organized_untransmitted_data['safe'].keys())
        intersection = unsafe_routes & safe_routes
        
        if intersection:
            print(f"âš ï¸  Data integrity violation: Routes {intersection} exist in both tabs!")
            return False
        return True
    
    def move_to_non_dups_duplicates(self, addr, disp_info):
        """Move a dispatch from dups to non-dups tab."""
        # Remove from dups tab
        if addr in self.organized_duplicates_data['dups']:
            if disp_info in self.organized_duplicates_data['dups'][addr]:
                self.organized_duplicates_data['dups'][addr].remove(disp_info)
                if not self.organized_duplicates_data['dups'][addr]:
                    del self.organized_duplicates_data['dups'][addr]
        
        # Ensure dispatch doesn't exist in dups anymore
        if addr in self.organized_duplicates_data['dups']:
            assert disp_info not in self.organized_duplicates_data['dups'][addr], f"Dispatch {disp_info} still exists in dups after move"
        
        # Add to non-dups tab
        if addr not in self.organized_duplicates_data['non_dups']:
            self.organized_duplicates_data['non_dups'][addr] = []
        self.organized_duplicates_data['non_dups'][addr].append(disp_info)
        
        self.refresh_duplicates_popup_content()
    
    def remove_from_dups_duplicates(self, addr, disp_info):
        """Remove a dispatch from dups tab."""
        if addr in self.organized_duplicates_data['dups']:
            self.organized_duplicates_data['dups'][addr].remove(disp_info)
            if not self.organized_duplicates_data['dups'][addr]:
                del self.organized_duplicates_data['dups'][addr]
            
            # Track this removed duplicate
            self.removed_duplicates.add((disp_info, addr))
            
            # Update the counter immediately
            self._update_duplicates_counter()
        
        self.refresh_duplicates_popup_content()
    
    def remove_from_non_dups_duplicates(self, addr, disp_info):
        """Remove a dispatch from non-dups tab."""
        if addr in self.organized_duplicates_data['non_dups']:
            self.organized_duplicates_data['non_dups'][addr].remove(disp_info)
            if not self.organized_duplicates_data['non_dups'][addr]:
                del self.organized_duplicates_data['non_dups'][addr]
            
            # Track this removed duplicate
            self.removed_duplicates.add((disp_info, addr))
            
            # Update the counter immediately
            self.refresh_duplicates_popup_content()
    
    def _update_duplicates_counter(self):
        """Update the duplicates counter to reflect current popup content"""
        try:
            # Calculate total duplicates from current organized data
            total_dups = sum(len(disp_list) for disp_list in self.organized_duplicates_data['dups'].values())
            
            # Update the counter display
            if hasattr(self, 'duplicates_counter'):
                self.duplicates_counter.show_text(str(total_dups))
                
            # Update the duplicates label if it exists (keep original text)
            if hasattr(self, 'duplicates_label'):
                self.duplicates_label.configure(text="DUPS")
                
        except Exception as e:
            logging.error(f"Error updating duplicates counter: {e}")
    
    def move_back_to_dups_duplicates(self, addr, disp_info):
        """Move a dispatch from non-dups back to dups tab."""
        # Remove from non-dups tab
        if addr in self.organized_duplicates_data['non_dups']:
            if disp_info in self.organized_duplicates_data['non_dups'][addr]:
                self.organized_duplicates_data['non_dups'][addr].remove(disp_info)
                if not self.organized_duplicates_data['non_dups'][addr]:
                    del self.organized_duplicates_data['non_dups'][addr]
        
        # Ensure dispatch doesn't exist in non-dups anymore
        if addr in self.organized_duplicates_data['non_dups']:
            assert disp_info not in self.organized_duplicates_data['non_dups'][addr], f"Dispatch {disp_info} still exists in non-dups after move"
        
        # Add to dups tab
        if addr not in self.organized_duplicates_data['dups']:
            self.organized_duplicates_data['dups'][addr] = []
        self.organized_duplicates_data['dups'][addr].append(disp_info)
        
        self.refresh_duplicates_popup_content()
    
    def verify_duplicates_data_integrity(self):
        """Verify that no dispatches exist in both dups and non-dups tabs simultaneously."""
        all_dups = set()
        all_non_dups = set()
        
        for addr, disp_list in self.organized_duplicates_data['dups'].items():
            for disp_info in disp_list:
                all_dups.add((addr, disp_info))
        
        for addr, disp_list in self.organized_duplicates_data['non_dups'].items():
            for disp_info in disp_list:
                all_non_dups.add((addr, disp_info))
        
        intersection = all_dups & all_non_dups
        
        if intersection:
            print(f"âš ï¸  Duplicates data integrity violation: Dispatches {intersection} exist in both tabs!")
            return False
        return True

    def show_duplicates_popup(self):
        """Show popup with duplicate routes information with tabs for organization."""
        try:
            # Check if we're in widget context
            if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
                parent_window = self.widget_window
            else:
                parent_window = self
            
            if not hasattr(self, 'duplicates_popup_data') or not self.duplicates_popup_data:
                messagebox.showinfo("Duplicate Stops", "No duplicate stops found.")
                return

            # Initialize organized data if not exists
            if not hasattr(self, 'organized_duplicates_data'):
                self.organized_duplicates_data = {
                    'dups': self.duplicates_popup_data.copy(),
                    'non_dups': {}
                }
            else:
                # Preserve existing organization, only add NEW duplicates that don't exist in either tab
                # This prevents overwriting your manual organization
                new_dups = {}
                for addr, disp_list in self.duplicates_popup_data.items():
                    # Check if this address exists in non_dups
                    non_dup_dispatches = set()
                    if addr in self.organized_duplicates_data['non_dups']:
                        non_dup_dispatches = set(self.organized_duplicates_data['non_dups'][addr])
                    
                    # Check if this address exists in current dups
                    existing_dup_dispatches = set()
                    if addr in self.organized_duplicates_data['dups']:
                        existing_dup_dispatches = set(self.organized_duplicates_data['dups'][addr])
                    
                    # Only add dispatches that are NEW (not in non_dups and not already in dups)
                    new_dispatches = [disp for disp in disp_list 
                                    if disp not in non_dup_dispatches 
                                    and disp not in existing_dup_dispatches]
                    
                    if new_dispatches:
                        new_dups[addr] = new_dispatches
                
                # Merge with existing dups (preserve your organization)
                for addr, disp_list in new_dups.items():
                    if addr in self.organized_duplicates_data['dups']:
                        self.organized_duplicates_data['dups'][addr].extend(disp_list)
                    else:
                        self.organized_duplicates_data['dups'][addr] = disp_list.copy()
                
                # Remove dispatches that are no longer active (no longer in source data)
                # This cleans up old entries without affecting your organization
                all_source_dispatches = set()
                for addr, disp_list in self.duplicates_popup_data.items():
                    for disp in disp_list:
                        all_source_dispatches.add((addr, disp))
                
                # Clean up dups tab
                for addr in list(self.organized_duplicates_data['dups'].keys()):
                    self.organized_duplicates_data['dups'][addr] = [
                        disp for disp in self.organized_duplicates_data['dups'][addr]
                        if (addr, disp) in all_source_dispatches
                    ]
                    if not self.organized_duplicates_data['dups'][addr]:
                        del self.organized_duplicates_data['dups'][addr]
                
                # Clean up non_dups tab
                for addr in list(self.organized_duplicates_data['non_dups'].keys()):
                    self.organized_duplicates_data['non_dups'][addr] = [
                        disp for disp in self.organized_duplicates_data['non_dups'][addr]
                        if (addr, disp) in all_source_dispatches
                    ]
                    if not self.organized_duplicates_data['non_dups'][addr]:
                        del self.organized_duplicates_data['non_dups'][addr]

            popup = ctk.CTkToplevel(parent_window)
            popup.overrideredirect(True)
            popup.grab_set()
            # Make popup appear above all windows including ROUTR
            popup.attributes("-topmost", True)
            popup.title("Duplicate Stops")
            
            outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
            outer.pack(fill="both", expand=True)
            inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
            inner.pack(fill="both", expand=True, padx=2, pady=2)

            header = ctk.CTkFrame(inner, fg_color="black")
            header.pack(fill="x")
            
            ctk.CTkLabel(header, text="Duplicate Stops",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="#FF6600").pack(side="left", padx=10, pady=5)
            
            def close_popup():
                popup.destroy()
                
            def copy_to_clipboard():
                lines = []
                lines.append("=== DUPLICATE STOPS ===")
                
                # Add dups items
                for addr, disp_list in self.organized_duplicates_data['dups'].items():
                    lines.append(f"âš  Address: {addr}")
                    for disp_info in disp_list:
                        lines.append(f"  {disp_info}")
                    lines.append("")
                
                # Add non-dups items
                for addr, disp_list in self.organized_duplicates_data['non_dups'].items():
                    lines.append(f"âœ“ Address: {addr}")
                    for disp_info in disp_list:
                        lines.append(f"  {disp_info}")
                    lines.append("")
                    
                copy_text = "\n".join(lines)
                popup.clipboard_clear()
                popup.clipboard_append(copy_text)
                popup.update()
            
            ctk.CTkButton(header, text="âœ•", width=30, height=25,
                      fg_color="#FF6600", hover_color="#800080",
                      command=close_popup).pack(side="right", padx=(5,10), pady=5)
            ctk.CTkButton(header, text="Copy", width=50, height=25,
                      fg_color="#FF6600", hover_color="#800080",
                      command=copy_to_clipboard).pack(side="right", padx=0, pady=5)
            
            def start_move(event):
                popup.x0 = event.x; popup.y0 = event.y
            def do_move(event):
                x = popup.winfo_x() + (event.x - popup.x0)
                y = popup.winfo_y() + (event.y - popup.y0)
                popup.geometry(f"+{x}+{y}")
            header.bind("<ButtonPress-1>", start_move)
            header.bind("<B1-Motion>", do_move)

            # Create single content area
            content = ctk.CTkScrollableFrame(inner, fg_color="black", height=300, width=400)
            content.pack(fill="both", expand=True, padx=10, pady=(0,10))
            
            # Store references for refresh functionality
            self.duplicates_popup_content = content
            

            
            # Populate all duplicates in single view
            all_items = []
            
            # Add dups items
            for addr, disp_list in self.organized_duplicates_data['dups'].items():
                for disp_info in disp_list:
                    all_items.append((addr, disp_info, "dups"))
            
            # Add non-dups items
            for addr, disp_list in self.organized_duplicates_data['non_dups'].items():
                for disp_info in disp_list:
                    all_items.append((addr, disp_info, "non_dups"))
            
            # Display all items
            for addr, disp_info, status in all_items[:100]:  # Limit to first 100
                # Display address header if it's the first occurrence
                if addr not in [item[0] for item in all_items[:all_items.index((addr, disp_info, status))]]:
                    ctk.CTkLabel(
                        content,
                        text=f"Address: {addr}",
                        font=ctk.CTkFont(size=12, weight="bold"),
                        text_color="#00AA00" if status == "non_dups" else "#FF6600"
                    ).pack(fill="x", anchor="w", padx=10, pady=(2,0))
                
                # Display dispatch info
                disp_frame = ctk.CTkFrame(content, fg_color="black")
                disp_frame.pack(fill="x", padx=10, pady=(0,2))
                
                # Status indicator and dispatch info
                status_color = "#00AA00" if status == "non_dups" else "white"
                status_text = "âœ“" if status == "non_dups" else "âš "
                
                ctk.CTkLabel(
                    disp_frame,
                    text=f"{status_text} {disp_info}",
                    font=ctk.CTkFont(size=12),
                    text_color=status_color
                ).pack(side="left", fill="x", expand=True, anchor="w", padx=(10,0))

            if len(all_items) > 100:
                ctk.CTkLabel(
                    content,
                    text=f"... and {len(all_items) - 100} more items",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))

            self._position_popup(popup)
            
            # Store popup reference for refresh functionality
            self.duplicates_popup = popup
        except Exception as e:
            logging.error(f"Error showing duplicates popup: {e}")
            import traceback
            traceback.print_exc()
            try:
                messagebox.showerror("Error", f"Failed to show duplicates popup: {e}")
            except:
                pass

    def refresh_duplicates_popup_content(self):
        """Refresh the duplicates popup content in real-time without closing"""
        try:
            if not hasattr(self, 'duplicates_popup') or not self.duplicates_popup:
                return
            
            # Update header counts
            total_dups = sum(len(disp_list) for disp_list in self.organized_duplicates_data['dups'].values())
            total_non_dups = sum(len(disp_list) for disp_list in self.organized_duplicates_data['non_dups'].values())
            
            # Find and update the header label
            for widget in self.duplicates_popup.winfo_children():
                if hasattr(widget, 'winfo_children'):
                    for child in widget.winfo_children():
                        if hasattr(child, 'winfo_children'):
                            for grandchild in child.winfo_children():
                                if isinstance(grandchild, ctk.CTkLabel) and "Duplicate Stops" in grandchild.cget("text"):
                                    grandchild.configure(text=f"Duplicate Stops (Dups: {total_dups}, Non-Dups: {total_non_dups})")
                                    break
            
            # Only add NEW duplicates that don't exist in either tab
            # This preserves your manual organization
            current_dups = {}
            for addr, disp_list in self.duplicates_popup_data.items():
                # Check if this address exists in non_dups
                non_dup_dispatches = set()
                if addr in self.organized_duplicates_data['non_dups']:
                    non_dup_dispatches = set(self.organized_duplicates_data['non_dups'][addr])
                
                # Check if this address exists in current dups
                existing_dup_dispatches = set()
                if addr in self.organized_duplicates_data['dups']:
                    existing_dup_dispatches = set(self.organized_duplicates_data['dups'][addr])
                
                # Only add dispatches that are NEW (not in non_dups and not already in dups)
                new_dispatches = [disp for disp in disp_list 
                                if disp not in non_dup_dispatches 
                                and disp not in existing_dup_dispatches]
                
                if new_dispatches:
                    current_dups[addr] = new_dispatches
            
            # Merge with existing dups (preserve your organization)
            for addr, disp_list in current_dups.items():
                if addr in self.organized_duplicates_data['dups']:
                    self.organized_duplicates_data['dups'][addr].extend(disp_list)
                else:
                    self.organized_duplicates_data['dups'][addr] = disp_list.copy()
            
            # Remove dispatches that are no longer active (no longer in source data)
            # This cleans up old entries without affecting your organization
            all_source_dispatches = set()
            for addr, disp_list in self.duplicates_popup_data.items():
                for disp in disp_list:
                    all_source_dispatches.add((addr, disp))
            
            # Clean up dups tab
            for addr in list(self.organized_duplicates_data['dups'].keys()):
                self.organized_duplicates_data['dups'][addr] = [
                    disp for disp in self.organized_duplicates_data['dups'][addr]
                    if (addr, disp) in all_source_dispatches
                ]
                if not self.organized_duplicates_data['dups'][addr]:
                    del self.organized_duplicates_data['dups'][addr]
            
            # Clean up non_dups tab
            for addr in list(self.organized_duplicates_data['non_dups'].keys()):
                self.organized_duplicates_data['non_dups'][addr] = [
                    disp for disp in self.organized_duplicates_data['non_dups'][addr]
                    if (addr, disp) in all_source_dispatches
                ]
                if not self.organized_duplicates_data['non_dups'][addr]:
                    del self.organized_duplicates_data['non_dups'][addr]
            
            # Clear and repopulate single content area
            for widget in self.duplicates_popup_content.winfo_children():
                widget.destroy()
            
            # Populate all duplicates in single view
            all_items = []
            
            # Add dups items
            for addr, disp_list in self.organized_duplicates_data['dups'].items():
                for disp_info in disp_list:
                    all_items.append((addr, disp_info, "dups"))
            
            # Add non-dups items
            for addr, disp_list in self.organized_duplicates_data['non_dups'].items():
                for disp_info in disp_list:
                    all_items.append((addr, disp_info, "non_dups"))
            
            # Display all items
            for addr, disp_info, status in all_items[:100]:  # Limit to first 100
                # Display address header if it's the first occurrence
                if addr not in [item[0] for item in all_items[:all_items.index((addr, disp_info, status))]]:
                    ctk.CTkLabel(
                        self.duplicates_popup_content,
                        text=f"Address: {addr}",
                        font=ctk.CTkFont(size=12, weight="bold"),
                        text_color="#00AA00" if status == "non_dups" else "#FF6600"
                    ).pack(fill="x", anchor="w", padx=10, pady=(2,0))
                
                # Display dispatch info
                disp_frame = ctk.CTkFrame(self.duplicates_popup_content, fg_color="black")
                disp_frame.pack(fill="x", padx=10, pady=(0,2))
                
                # Status indicator and dispatch info
                status_color = "#00AA00" if status == "non_dups" else "white"
                status_text = "âœ“" if status == "non_dups" else "âš "
                
                ctk.CTkLabel(
                    disp_frame,
                    text=f"{status_text} {disp_info}",
                    font=ctk.CTkFont(size=12),
                    text_color=status_color
                ).pack(side="left", fill="x", expand=True, anchor="w", padx=(10,0))

            if len(all_items) > 100:
                ctk.CTkLabel(
                    self.duplicates_popup_content,
                    text=f"... and {len(all_items) - 100} more items",
                    font=ctk.CTkFont(size=12),
                    text_color="#FF6600"
                ).pack(fill="x", anchor="w", padx=10, pady=(0,0))
            
            # Verify data integrity after refresh
            self.verify_duplicates_data_integrity()
                                    
        except Exception as e:
            print(f"ðŸ” Error refreshing duplicates popup: {e}")

    def _position_popup(self, popup):
        """Position popup in center of main window."""
        self.update_idletasks()
        mx, my = self.winfo_x(), self.winfo_y()
        mw, mh = self.winfo_width(), self.winfo_height()
        popup.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        x = mx + (mw - pw)//2; y = my + (mh - ph)//2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")

    def _convert_to_uppercase_sof(self, event=None):
        txt = self.sof_location_entry.get().upper()[:4]
        self.sof_location_entry.delete(0, ctk.END)
        self.sof_location_entry.insert(0, txt)  

    def load_sof_routes(self):
        # 0) Check if SOF+ monitor is already running
        if self.is_sof_monitor_running():
            messagebox.showwarning("SOF+ Monitor Running", 
                "SOF+ monitor is already running. Please wait for it to complete or reset the tab first.")
            return
            
        # 1) Read and validate the location
        loc = self.sof_location_entry.get().strip().upper()
        if not loc:
            messagebox.showerror("Missing Info", "Enter a dispatch location.")
            return
        
        # reset retry count on fresh load
        self.sof_load_attempts = 0

        # Store for the monitor thread
        self.sof_loc = loc

        # â”€â”€â”€ Disable inputs so you can't spam "Load" â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.sof_location_entry.configure(state="disabled")
        self.sof_load_button.configure(state="disabled")
        self.sof_reset_button.configure(state="disabled")

        # â”€â”€â”€ Show & start the inline loading bar *first* â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.sof_progress.pack(pady=(5, 0), padx=10, fill="x")
        self.sof_progress.start()

        # â”€â”€â”€ Then show the embedded SoF+ grid (below the bar) â”€â”€â”€â”€â”€â”€â”€â”€
        self.sof_grid_frame.pack(padx=(10,0), pady=(10,10), fill="both", expand=True)
        
        # â”€â”€â”€ Clear any old route buttons â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        for w in self.sof_grid_frame.winfo_children():
            w.destroy()

        def worker():
            # 1) Launch headless Chrome with matching chromedriver
            driver_path = ChromeDriverManager().install()
            service     = Service(driver_path)
            opts        = Options()
            opts.add_argument("--disable-extensions")
            opts.add_argument("--headless")  # Enabled headless mode for route loading
            opts.add_argument("--disable-gpu")
            opts.add_argument("--window-size=1920,1080")
            driver = webdriver.Chrome(service=service, options=opts)
            wait   = WebDriverWait(driver, 5)
            routes = []
            login_url = "https://internal.example.com"

            try:
                # 2) Go to SoF+ UI; if redirected away, perform login with MFA
                import threading
                # use a helper monitor instance to get the same MFA pop-ups
                login_helper = CarrierSeleniumMonitor(
                    username=self.username,
                    password=self.password,
                    dispatch_location=self.sof_loc,
                    stop_event=threading.Event(),
                    message_types=[],  # not used here
                    mfa_factor_callback=self.prompt_mfa_factor,
                    mfa_code_callback  =self.prompt_mfa_code
                )
                login_helper.sso_navigate(
                    driver, wait,
                    login_url,
                    self.username, self.password
                )

                # FAST home-page detection with up to 5 retries
                for attempt in range(5):
                    for _ in range(20):
                        if "/home" in driver.current_url:
                            break
                        time.sleep(0.1)
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH,
                                "//span[normalize-space(text())='FRO']"
                            ))
                        )
                        logging.info(f"âœ… Home UI detected on attempt {attempt+1}")
                        break
                    except TimeoutException:
                        logging.warning(f"â— SoF+ load home detection attempt {attempt+1} failed")
                        if attempt < 4:
                            logging.info("ðŸ”„ Refreshing and retryingâ€¦")
                            driver.refresh()
                            time.sleep(7)
                        else:
                            raise TimeoutException("SoF+: could not detect home UI after 5 attempts")

                # 3) Go to Status of Fleet via FRO menu
                fro_el = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']")))
                driver.execute_script("arguments[0].click();", fro_el)
                sof_el = wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//span[normalize-space(text())='Status of Fleet']"
                )))
                driver.execute_script("arguments[0].click();", sof_el)

                # 4) Enter dispatch location
                fld = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']"
                )))
                fld.clear()
                fld.send_keys(self.sof_loc, Keys.RETURN)

                # 5) Wait for the column filter input to appear
                header_input = wait.until(EC.presence_of_element_located((By.ID, "innerFilter")))
                header_th    = header_input.find_element(By.XPATH, "./ancestor::th")
                col_index    = len(header_th.find_elements(By.XPATH, "preceding-sibling::th")) + 1

                # 6) Grab every <td> in that column and collect numeric routes
                xpath_td = f"//table//tr/td[{col_index}]/div"
                wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_td)))
                elems  = driver.find_elements(By.XPATH, xpath_td)
                routes = [e.text.strip() for e in elems if e.text.strip().isdigit()]
                
                # Check if no routes were found
                if not routes:
                    raise Exception(f"No routes found for station ID: {self.sof_loc}")

            except TimeoutException:
                logging.error("âŒ SoF+ login or navigation timed out; aborting load")
                self.after(0, lambda: on_ui_error("Unable to authenticate or load SoF+"))
                return

            finally:
                driver.quit()

            # 7) On the UI thread: tear down the bar and build buttons
            def on_ui():
                # â”€â”€â”€ Tear down inline loading bar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                self.sof_progress.stop()
                self.sof_progress.pack_forget()

                # â”€â”€â”€ Enable Reset (lock Load/entry until user clicks Reset) â”€
                self.sof_reset_button.configure(state="normal")

                # â”€â”€â”€ Build route buttons â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                self.sof_route_buttons = {}
                route_font = ctk.CTkFont(family="Tahoma", size=16, weight="bold")
                cols = 8

                for i, route in enumerate(routes):
                    btn = ctk.CTkButton(
                        self.sof_grid_frame,
                        text=route,
                        width=60, height=40,
                        font=route_font,
                        fg_color=BRAND_ORANGE,
                        text_color="white",
                        border_width=2,
                        border_color=BRAND_ORANGE
                    )
                    btn.grid(row=i // cols, column=i % cols, padx=5, pady=5)
                    self.sof_route_buttons[route] = btn
                    
                    # Scroll wheel binding will be handled by _bind_sof_scroll_wheel() after all buttons are created
                    
                    # Add tooltip for this route button
                    tooltip = CustomTooltip(btn, f"Route {route}\n\nLoading status...", delay=200)
                    # Store tooltip reference for later updates
                    if not hasattr(self, 'sof_route_tooltips'):
                        self.sof_route_tooltips = {}
                    self.sof_route_tooltips[route] = tooltip

                # Keep track of cleared state & start monitoring
                self.sof_route_cleared_states = {r: False for r in self.sof_route_buttons}
                # Initialize route colors dictionary
                self.sof_route_colors = {r: "#404040" for r in self.sof_route_buttons}
                
                # Make the scrollbar skinny to prevent cutting off route buttons
                self.sof_grid_frame._scrollbar.configure(width=8)
                
                # Add scroll wheel support for the SoF+ grid frame
                self._bind_sof_scroll_wheel_simple()
                
                # Start SOF+ monitor (checks each route 1 by 1)
                self.start_sof_monitor()
                
                # Start test function scheduler (runs every 15/45 minutes)
                self.start_test_scheduler()
                
                # Run test function immediately in background thread (with small delay to avoid overwhelming server)
                self.after(2000, lambda: threading.Thread(target=self._run_test_function_worker, daemon=True).start())
                
                # Show loading spinners in counters and status (only when test function starts)
                self.after(2000, lambda: self._show_loading_counters())
                self.after(2000, lambda: self.test_status_label.configure(text="ðŸ”„ Collecting pickup statistics..."))

            self.after(0, on_ui)

        # 8) Define an error callback for the UI thread
        def on_ui_error(msg):
            # stop & hide the loading bar
            self.sof_progress.stop()
            self.sof_progress.pack_forget()
            # re-enable Load UI
            self.sof_location_entry.configure(state="normal")
            self.sof_load_button   .configure(state="normal")
            
            # Check if this is likely a bad station ID error (no routes found)
            if "no routes" in msg.lower() or "timeout" in msg.lower() or "unable to authenticate" in msg.lower():
                # Don't retry for clear user input errors - let user fix the station ID
                messagebox.showerror(
                    "SoF+ Load Failed",
                    f"{msg}\n\nPlease check the station ID and try again."
                )
                return
            
            # increment and decide whether to retry for other errors
            self.sof_load_attempts += 1
            if self.sof_load_attempts < 5:
                messagebox.showwarning(
                    "SoF+ Load Error",
                    f"{msg}\nAutomatically retrying ({self.sof_load_attempts}/5)â€¦"
                )
                # retry after a short pause
                self.after(1000, self.load_sof_routes)
            else:
                # give up after 5 tries
                messagebox.showerror(
                    "SoF+ Load Failed",
                    "Could not load SoF+ after 5 attempts.\n"
                    "Please check network/credentials and press Load again."
                )

        # 9) Spin off the worker thread
        def _worker_wrapper():
            try:
                worker()
            except Exception as e:
                logging.exception("âŒ SoF+ load thread crashed")
                # properly pass the exception message into the UI error callback
                self.after(0, on_ui_error, str(e))

        threading.Thread(target=_worker_wrapper, daemon=True).start()
    def start_sof_monitor(self):
        # Prevent multiple monitors from running simultaneously
        if hasattr(self, "sof_monitor_running") and self.sof_monitor_running:
            logging.info("âš ï¸ SOF+ monitor already running, skipping start request")
            return
            
        # Stop any existing monitor
        if hasattr(self, "sof_monitor_stop_event"):
            self.sof_monitor_stop_event.set()
            # Wait a moment for the thread to stop
            time.sleep(0.5)

        # Create a fresh stop event
        self.sof_monitor_stop_event = threading.Event()
        
        # Set the running flag
        self.sof_monitor_running = True

        # â”€â”€â”€ Show monitor-startup bar *above* the grid â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # 1) temporarily un-pack the grid so it won't cover the bar
        self.sof_grid_frame.pack_forget()

        # 2) pack & start the bar
        self.sof_monitor_progress.pack(pady=(5, 0), padx=10, fill="x")
        self.sof_monitor_progress.start()

        # 3) re-pack the grid so it sits below the bar
        self.sof_grid_frame.pack(padx=10, pady=(10,10), fill="both", expand=True)

        # Launch the monitoring thread
        t = threading.Thread(target=self._sof_monitor_worker, daemon=True)
        t.start()
        # record an initial "last update" so watchdog will notice if run() hangs
        self._last_sof_update = datetime.now()
        
        logging.info("âœ… SOF+ monitor started successfully")
        
    def stop_sof_monitor(self):
        """Stop the SOF+ monitor and reset the running flag."""
        if hasattr(self, "sof_monitor_stop_event"):
            self.sof_monitor_stop_event.set()
            # Wait a moment for the thread to stop
            time.sleep(0.5)
        
        if hasattr(self, "sof_monitor_running"):
            self.sof_monitor_running = False
            
        logging.info("ðŸ›‘ SOF+ monitor stopped")
        
    def is_sof_monitor_running(self):
        """Check if the SOF+ monitor is currently running."""
        return hasattr(self, "sof_monitor_running") and self.sof_monitor_running
        
    def _update_all_route_tooltips(self):
        """Update all route tooltips with current status information."""
        if hasattr(self, 'sof_route_tooltips'):
            for route, tooltip in self.sof_route_tooltips.items():
                try:
                    status_info = self._get_route_status_info(str(route))
                    tooltip.update_text(f"Route {route}\n\n{status_info}")
                except Exception as e:
                    # If status info not available yet, show loading message
                    tooltip.update_text(f"Route {route}\n\nLoading status...")
                    
    def _update_route_tooltip(self, route):
        """Update a single route tooltip with current status information."""
        if hasattr(self, 'sof_route_tooltips') and route in self.sof_route_tooltips:
            try:
                tooltip = self.sof_route_tooltips[route]
                status_info = self._get_route_status_info(route)
                tooltip.update_text(f"Route {route}\n\n{status_info}")
            except Exception as e:
                # If status info not available yet, show loading message
                tooltip.update_text(f"Route {route}\n\nLoading status...")



    def reset_sof_tab(self):
        """Clear SoF+ tab and re-enable Load for a new location."""
        # 1) Stop any running SoF+ monitor thread
        self.stop_sof_monitor()

        # 2) Clear out all route buttons & hide the grid
        for w in self.sof_grid_frame.winfo_children():
            w.destroy()
        self.sof_grid_frame.pack_forget()
        
        # Hide DEX 03 and PUX 03 counters when resetting
        if hasattr(self, 'dex03_frame'):
            self.dex03_frame.grid_remove()
        if hasattr(self, 'pux03_frame'):
            self.pux03_frame.grid_remove()

        # 3) Hide any progress bars
        self.sof_progress.stop()
        self.sof_progress.pack_forget()
        if hasattr(self, "sof_monitor_progress"):
            self.sof_monitor_progress.stop()
            self.sof_monitor_progress.pack_forget()

        # 4) Reset internal state dicts (optional but clean)
        self.sof_route_buttons = {}
        self.sof_route_states = {}
        self.sof_route_cleared_states = {}
        
        # 5) Re-enable Load & entry
        self.sof_location_entry.configure(state="normal")
        self.sof_load_button.configure(state="normal")

        # 6) Disable Reset until next load
        self.sof_reset_button.configure(state="disabled")
        
        # 7) Stop test function scheduler
        self.stop_test_scheduler()
        
        # 8) Reset counters to blank state
        self._reset_counters()
        


    def _get_route_status_info(self, route):
        """Get route status information based on legend colors and border colors."""
        try:
            # Get current route state from SoF+ monitor
            if hasattr(self, 'sof_route_states') and route in self.sof_route_states:
                state = self.sof_route_states[route]
                fill_color = state.get('chosen_color', '#404040')
                
                # Get border color from the actual button
                border_color = None
                if hasattr(self, 'sof_route_buttons') and route in self.sof_route_buttons:
                    try:
                        btn = self.sof_route_buttons[route]
                        border_color = btn.cget("border_color")
                        border_width = btn.cget("border_width")
                        # Only consider border if it has width > 0
                        if border_width == 0:
                            border_color = None
                    except:
                        pass
                
                # Build status message based on fill color
                if fill_color == "#FF0000":  # Red fill
                    status = "URGENT\nStops exist but route not locked/signed in"
                elif fill_color == "#006400":  # Green fill
                    status = "CLEARED\nRoute completed and cleared"
                elif fill_color == "#CCCC00":  # Yellow fill
                    status = "BREAK\nDriver on break"
                elif fill_color == "#404040":  # Dark Grey fill
                    status = "NORMAL\nRoute is active and locked"
                else:
                    status = f"Status: {fill_color}\nRoute information available"
                
                # Add border color information
                if border_color:
                    if border_color == "#FF0000":  # Red border
                        status += "\n\nBORDER: Stop closing within 30 minutes or past closing time"
                    elif border_color == "#FF6600":  # Orange border (BRAND_ORANGE)
                        status += "\n\nBORDER: Route has stops to pick up"
                
                return status
            
            else:
                # Fallback if state not available
                return "Status information not available\n\nRoute may still be loading or monitoring not started."
                
        except Exception as e:
            return f"Error getting status: {str(e)}"

        # 5) Re-enable Load & entry
        self.sof_location_entry.configure(state="normal")
        self.sof_load_button.configure(state="normal")

        # 6) Disable Reset until next load
        self.sof_reset_button.configure(state="disabled")
        
        # 7) Stop test function scheduler and SOF+ monitor
        self.stop_test_scheduler()
        if hasattr(self, 'sof_monitor_stop_event'):
            self.sof_monitor_stop_event.set()
        
        # 8) Reset counters to blank state
        self._reset_counters()    

    def show_legend_popup(self):
        """Pop-up showing the SoF+ color & highlight legend."""
        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.grab_set()

        # Outer border
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        # Inner content frame
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        # Header with title and close button
        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Legend",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(side="left", padx=10, pady=5)
        ctk.CTkButton(
            header, text="âœ•",
            width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=popup.destroy
        ).pack(side="right", padx=10, pady=5)

        # Content area
        content = ctk.CTkFrame(inner, fg_color="black")
        content.pack(fill="both", expand=True, padx=10, pady=(0,10))

        # Legend entries
        legend_items = [
            ("Grey fill",         "Route with stops assigned at some point"),
            ("Yellow fill",       "Route on Break"),
            ("Green fill",        "Cleared"),
            ("Red fill",          "Stops assigned/Problem with courier sign in"),
            ("Red border",        "Stop closing in less than 30 min or past closing time"),
            ("Orange border",     "Route has pick ups open"),
        ]
        for label, desc in legend_items:
            ctk.CTkLabel(
                content,
                text=f"â€¢ {label}: {desc}",
                font=ctk.CTkFont(size=12),
                text_color="white",
                anchor="w", justify="left"
            ).pack(fill="x", pady=2)

        # Center popup over main window
        self.update_idletasks()
        mx, my = self.winfo_x(), self.winfo_y()
        mw, mh = self.winfo_width(), self.winfo_height()
        popup.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        x = mx + (mw - pw)//2
        y = my + (mh - ph)//2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")
        
    def run_test_function(self):
        """Run the test function to collect pickup statistics using CSV data analysis."""
        try:
            
            # Create a headless Chrome driver with download capabilities
            driver_path = ChromeDriverManager().install()
            service = Service(driver_path, log_path=os.devnull)
            opts = Options()
            opts.add_argument("--disable-gpu")
            opts.add_argument("--headless")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--log-level=3")
            opts.add_experimental_option("excludeSwitches", ["enable-logging"])
            
            # Enable downloads
            prefs = {
                "download.default_directory": os.path.join(os.getcwd(), "downloads"),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            }
            opts.add_experimental_option("prefs", prefs)
            
            driver = webdriver.Chrome(service=service, options=opts)
            wait = WebDriverWait(driver, 20)
            
            # Create downloads directory
            download_dir = os.path.join(os.getcwd(), "downloads")
            os.makedirs(download_dir, exist_ok=True)
            
            try:
                # Login using the same method as other functions
                login_helper = CarrierSeleniumMonitor(
                    username=self.username,
                    password=self.password,
                    dispatch_location=self.sof_loc,
                    stop_event=threading.Event(),
                    message_types=[],
                    mfa_factor_callback=self.prompt_mfa_factor,
                    mfa_code_callback=self.prompt_mfa_code
                )
                
                login_helper.sso_navigate(
                    driver, wait,
                    "https://internal.example.com",
                    self.username, self.password
                )
                
                # Wait for home page
                for attempt in range(5):
                    for _ in range(20):
                        if "/home" in driver.current_url:
                            break
                        time.sleep(0.1)
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='FRO']"))
                        )
                        break
                    except TimeoutException:
                        if attempt < 4:
                            driver.refresh()
                            time.sleep(1)
                        else:
                            raise Exception("Failed to reach home page")
                
                # Navigate to Pickup List
                fro_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']")))
                fro_btn.click()
                time.sleep(0.5)
                
                pickup_list_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Pickup List']")))
                pickup_list_btn.click()
                time.sleep(1)
                
                # Set dispatch location
                fld = wait.until(EC.element_to_be_clickable((By.ID, "dispatchLocation")))
                fld.clear()
                fld.send_keys(self.sof_loc + Keys.RETURN)
                time.sleep(3)
                
                # Wait for table to load
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
                
                # Now use the proper CSV download and analysis approach
                # Run in separate thread to avoid blocking UI
                threading.Thread(target=self._run_test_function_worker, daemon=True).start()
                
            finally:
                driver.quit()
                
        except Exception as e:
            logging.error(f"Test function error: {e}")
    
    def _update_counters(self, una, early, late, untransmitted, duplicates, tooltip_data):
        """Update the counter labels with new values and popup data."""
        self.una_count = una
        self.early_count = early
        self.late_count = late
        self.untransmitted_count = untransmitted
        self.duplicates_count = duplicates
        
        # Store popup data
        self.una_popup_data = tooltip_data.get('una', [])
        self.early_popup_data = tooltip_data.get('early', [])
        self.late_popup_data = tooltip_data.get('late', [])
        self.untransmitted_popup_data = tooltip_data.get('untransmitted', {})
        self.duplicates_popup_data = tooltip_data.get('duplicates', {})
        
        # Initialize organized duplicates data if this is the first time or if new duplicates are found
        if not hasattr(self, 'organized_duplicates_data') or not self.organized_duplicates_data['dups']:
            self.organized_duplicates_data = {
                'dups': self.duplicates_popup_data.copy(),
                'non_dups': {}
            }
        else:
            # Check for returned duplicates (stops that were removed but are back in new data)
            returned_duplicates = set()
            for addr, disp_list in self.duplicates_popup_data.items():
                for disp_info in disp_list:
                    if (disp_info, addr) in self.removed_duplicates:
                        returned_duplicates.add((disp_info, addr))
                        logging.info(f"ðŸ”„ Duplicate returned: {disp_info} at {addr}")
            
            # Remove returned duplicates from the removed_duplicates set
            self.removed_duplicates -= returned_duplicates
            
            # Add any new duplicates to the existing organized data
            for addr, disp_list in self.duplicates_popup_data.items():
                if addr not in self.organized_duplicates_data['dups']:
                    self.organized_duplicates_data['dups'][addr] = disp_list.copy()
                else:
                    # Add any new dispatch numbers that weren't already there
                    existing_disps = set(self.organized_duplicates_data['dups'][addr])
                    for disp_info in disp_list:
                        if disp_info not in existing_disps:
                            self.organized_duplicates_data['dups'][addr].append(disp_info)
        
        # Initialize organized untransmitted data if this is the first time or if new untransmitted data is found
        if not hasattr(self, 'organized_untransmitted_data') or not self.organized_untransmitted_data['unsafe']:
            self.organized_untransmitted_data = {
                'unsafe': self.untransmitted_popup_data.copy(),
                'safe': {}
            }
        else:
            # Add any new untransmitted routes to the existing organized data
            for route, count in self.untransmitted_popup_data.items():
                if route not in self.organized_untransmitted_data['unsafe']:
                    self.organized_untransmitted_data['unsafe'][route] = count
        
        self.una_counter.show_text(una)
        self.early_counter.show_text(early)
        self.late_counter.show_text(late)
        self.untransmitted_counter.show_text(untransmitted)
        self.duplicates_counter.show_text(duplicates)
        
        # Hide status label after successful update
        self.test_status_label.pack_forget()
        
        # Update widget content if widget is visible
        if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
            self.update_widget_content()
    

    
    def _show_loading_counters(self):
        """Show loading spinners in all counter labels."""
        # Show the labels first
        self.una_label.pack()
        self.early_label.pack()
        self.late_label.pack()
        self.untransmitted_label.pack()
        self.duplicates_label.pack()
        # Ensure DEX/PUX frames are visible alongside the main counters
        if hasattr(self, "dex03_frame"):
            self.dex03_frame.grid()
        if hasattr(self, "pux03_frame"):
            self.pux03_frame.grid()
        
        # Then show the spinners
        self.una_counter.show_spinner()
        self.early_counter.show_spinner()
        self.late_counter.show_spinner()
        self.untransmitted_counter.show_spinner()
        self.duplicates_counter.show_spinner()
    
    def _reset_counters(self):
        """Reset counters to blank state and hide labels."""
        # Reset counter values
        self.una_count = 0
        self.early_count = 0
        self.late_count = 0
        self.untransmitted_count = 0
        self.duplicates_count = 0
        
        # Hide all labels
        self.una_label.pack_forget()
        self.early_label.pack_forget()
        self.late_label.pack_forget()
        self.untransmitted_label.pack_forget()
        self.duplicates_label.pack_forget()
        # Hide DEX/PUX frames during reset so they don't show on next load
        if hasattr(self, "dex03_frame"):
            self.dex03_frame.grid_remove()
        if hasattr(self, "pux03_frame"):
            self.pux03_frame.grid_remove()
        
        # Reset counters to blank
        self.una_counter.show_text("")
        self.early_counter.show_text("")
        self.late_counter.show_text("")
        self.untransmitted_counter.show_text("")
        self.duplicates_counter.show_text("")
        
        # Reset organized duplicates data
        self.organized_duplicates_data = {
            'dups': {},
            'non_dups': {}
        }
        
        # Reset organized untransmitted data
        self.organized_untransmitted_data = {
            'unsafe': {},
            'safe': {}
        }
        
        # Clear removed duplicates tracking
        self.removed_duplicates.clear()
        
        # Update widget content if widget is visible
        if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
            self.update_widget_content()
        
        # Hide status label
        self.test_status_label.pack_forget()
    
    def _run_test_function_worker(self):
        """Worker thread for running the test function without blocking UI."""
        # Show loading spinners in counters and update status
        self.after(0, lambda: self._show_loading_counters())
        self.after(0, lambda: self.test_status_label.configure(text="ðŸ”„ Collecting pickup statistics..."))
        
        try:
            # Create a headless Chrome driver with download capabilities
            driver_path = ChromeDriverManager().install()
            service = Service(driver_path, log_path=os.devnull)
            opts = Options()
            opts.add_argument("--disable-gpu")
            opts.add_argument("--headless")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--log-level=3")
            opts.add_experimental_option("excludeSwitches", ["enable-logging"])
            
            # Enable downloads
            prefs = {
                "download.default_directory": os.path.join(os.getcwd(), "downloads"),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            }
            opts.add_experimental_option("prefs", prefs)
            
            driver = webdriver.Chrome(service=service, options=opts)
            wait = WebDriverWait(driver, 20)
            
            # Create downloads directory
            download_dir = os.path.join(os.getcwd(), "downloads")
            os.makedirs(download_dir, exist_ok=True)
            
            try:
                # Login using the same method as other functions
                login_helper = CarrierSeleniumMonitor(
                    username=self.username,
                    password=self.password,
                    dispatch_location=self.sof_loc,
                    stop_event=threading.Event(),
                    message_types=[],
                    mfa_factor_callback=self.prompt_mfa_factor,
                    mfa_code_callback=self.prompt_mfa_code
                )
                
                login_helper.sso_navigate(
                    driver, wait,
                    "https://internal.example.com",
                    self.username, self.password
                )
                
                # Wait for home page
                for attempt in range(5):
                    for _ in range(20):
                        if "/home" in driver.current_url:
                            break
                        time.sleep(0.1)
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='FRO']"))
                        )
                        break
                    except TimeoutException:
                        if attempt < 4:
                            driver.refresh()
                            time.sleep(1)
                        else:
                            raise Exception("Failed to reach home page")
                
                # Navigate to Pickup List
                fro_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']")))
                fro_btn.click()
                time.sleep(0.5)
                
                pickup_list_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Pickup List']")))
                pickup_list_btn.click()
                time.sleep(1)
                
                # Set dispatch location
                fld = wait.until(EC.element_to_be_clickable((By.ID, "dispatchLocation")))
                fld.clear()
                fld.send_keys(self.sof_loc + Keys.RETURN)
                time.sleep(3)
                
                # Wait for table to load
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
                
                # Now use the proper CSV download and analysis approach
                una_count, early_count, late_count, untransmitted_count, duplicates_count, tooltip_data = self._get_pickup_statistics_from_csv(driver, wait, download_dir)
                
                # Update counters on UI thread
                self.after(0, self._update_counters, una_count, early_count, late_count, untransmitted_count, duplicates_count, tooltip_data)
                
                # Update status to show completion
                self.after(0, lambda: self.test_status_label.configure(text="âœ… Pickup statistics updated"))
                
            finally:
                driver.quit()
                
        except Exception as e:
            logging.error(f"Test function worker error: {e}")
    
    def _get_pickup_statistics_from_csv(self, driver, wait, download_dir):
        """Get pickup statistics by downloading and analyzing CSV data."""
        try:
            # First, open multiselect dropdown and select all items
            self._open_multiselect_dropdown(driver, wait)
            self._select_all_in_multiselect(driver, wait)
            
            # Download CSV
            csv_path = self._download_csv(driver, wait, download_dir)
            
            # Read and analyze CSV data
            if csv_path and os.path.exists(csv_path):
                una_count, early_count, late_count, untransmitted_count, duplicates_count, tooltip_data = self._analyze_csv_data(csv_path)
                
                # Clean up downloaded file
                try:
                    os.remove(csv_path)
                except:
                    pass
                
                return una_count, early_count, late_count, untransmitted_count, duplicates_count, tooltip_data
            else:
                logging.error("CSV download failed")
                empty_tooltip_data = {
                    'una': [],
                    'early': [],
                    'late': [],
                    'untransmitted': {},
                    'duplicates': {}
                }
                return 0, 0, 0, 0, 0, empty_tooltip_data
                
        except Exception as e:
            logging.error(f"Error getting pickup statistics from CSV: {e}")
            empty_tooltip_data = {
                'una': [],
                'early': [],
                'late': [],
                'untransmitted': {},
                'duplicates': {}
            }
            return 0, 0, 0, 0, 0, empty_tooltip_data
    
    def _open_multiselect_dropdown(self, driver, wait):
        """Open the multiselect dropdown using the same approach as the working test script."""
        root, frame_idx = self._find_element_across_frames(driver, (By.CSS_SELECTOR, "div.p-multiselect"))
        if frame_idx is not None:
            driver.switch_to.frame(frame_idx)
        
        target = self._get_visible_multiselect_trigger(driver)
        if not target:
            self._switch_to_default(driver)
            el, frame_idx = self._find_element_across_frames(driver, (By.XPATH, "//path[starts-with(@d,'M7.01744 10.398')]/ancestor::*[self::button or self::*[contains(@class,'p-multiselect')]][1]"))
            if el:
                if frame_idx is not None:
                    driver.switch_to.frame(frame_idx)
                target = el
        
        if not target:
            raise TimeoutException("No visible MultiSelect trigger found (checked default & iframes).")
        
        self._safe_click(driver, target)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.p-multiselect-panel")))
        logging.info("âœ… Opened multiselect dropdown")
    
    def _get_visible_multiselect_trigger(self, driver):
        """Get visible multiselect trigger using the same approach as the working test script."""
        try:
            ms_list = driver.find_elements(By.CSS_SELECTOR, "div.p-multiselect")
            for ms in ms_list:
                btn = None
                try:
                    btn = ms.find_element(By.CSS_SELECTOR, "button.p-multiselect-trigger")
                except:
                    pass
                target = btn or ms
                visible = driver.execute_script(
                    """
                    const el = arguments[0];
                    if (!el) return false;
                    const st = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    return st.visibility !== 'hidden' && st.display !== 'none' && r.width > 0 && r.height > 0;
                    """,
                    target,
                )
                if visible:
                    return target
        except Exception:
            pass
        return None
    
    def _select_all_in_multiselect(self, driver, wait):
        """Select all items in the multiselect dropdown."""
        try:
            # Find the select all checkbox
            select_all_cb = wait.until(EC.presence_of_element_located((
                By.XPATH, "//div[contains(@class,'p-multiselect-panel')]//div[contains(@class,'p-multiselect-header')]//div[@role='checkbox']"
            )))
            
            # Check if already selected
            aria_checked = (select_all_cb.get_attribute("aria-checked") or "").lower()
            if aria_checked not in ("true", "mixed"):
                select_all_cb.click()
                wait.until(lambda d: ((select_all_cb.get_attribute("aria-checked") or "").lower() in ("true", "mixed")))
                logging.info("âœ… Selected all items in dropdown")
            else:
                logging.info("â„¹ï¸ 'Select all' already active")
        except Exception as e:
            logging.error(f"Error selecting all items: {e}")
    
    def _download_csv(self, driver, wait, download_dir):
        """Download CSV and return the file path using the same method as the working test script."""
        try:
            click_start = time.time()
            
            # Find the download button using the same approach as the working script
            el, frame_idx = self._find_element_across_frames(driver, (By.ID, "download_cvs"))
            if not el:
                raise TimeoutException("Could not find 'Download - CSV' button.")
            
            if frame_idx is not None:
                driver.switch_to.frame(frame_idx)
            
            # Use the exact same button finding approach
            download_btn = wait.until(EC.element_to_be_clickable((By.ID, "download_cvs")))
            self._safe_click(driver, download_btn)
            logging.info("âœ… Clicked download CSV button")
            
            # Wait for download to complete using the same timing
            csv_path = self._wait_for_csv_download(download_dir, click_start, timeout=120)
            if csv_path:
                logging.info(f"âœ… CSV download complete: {csv_path}")
                return csv_path
            else:
                logging.error("âŒ CSV download timed out")
                return None
                
        except Exception as e:
            logging.error(f"Error downloading CSV: {e}")
            return None
    
    def _wait_for_csv_download(self, download_dir, since_ts, timeout=120):
        """Wait for CSV download to complete using the same method as the working test script."""
        import glob
        end = time.time() + timeout
        last_size = -1
        last_path = None
        
        while time.time() < end:
            csvs = [p for p in glob.glob(os.path.join(download_dir, "*.csv")) if os.path.getmtime(p) >= since_ts - 1]
            if csvs:
                csvs.sort(key=os.path.getmtime, reverse=True)
                candidate = csvs[0]
                partials = glob.glob(candidate + ".crdownload")
                if not partials:
                    size = os.path.getsize(candidate)
                    if candidate == last_path and size == last_size:
                        return candidate
                    last_path, last_size = candidate, size
            time.sleep(0.5)
        
        raise TimeoutException("Timed out waiting for CSV download to finish.")
    
    def _analyze_csv_data(self, csv_path):
        """Analyze CSV data using the same method as the working test script."""
        try:
            # Read CSV with the same robust method as the working script
            import csv
            import io
            
            with open(csv_path, "rb") as f:
                raw = f.read()
            
            # Use same encoding detection as working script
            encoding_used = "utf-8"
            text = ""
            for enc in ("utf-8-sig", "utf-8", "latin-1"):
                try:
                    text = raw.decode(enc)
                    encoding_used = enc
                    break
                except Exception:
                    continue
            
            # Use same delimiter detection as working script
            try:
                first_line = text.splitlines()[0] if text else ""
                dialect = csv.Sniffer().sniff(first_line or ",")
                sep = dialect.delimiter
            except Exception:
                sep = "," if text.count(",") >= text.count(";") else ";"
            
            df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str, engine="python", on_bad_lines="skip")
            logging.info(f"ðŸ“Š CSV loaded: {len(df)} rows, {len(df.columns)} columns")
            logging.info(f"ðŸ“Š Columns: {list(df.columns)}")
            
            # Use the exact same column mapping as the working script
            colmap = {c.lower().strip(): c for c in df.columns}
            
            def _pick_column(colmap, candidates):
                for key in candidates:
                    if key in colmap:
                        return colmap[key]
                normalized = {re.sub(r"[^a-z0-9]", "", k): v for k, v in colmap.items()}
                for key in candidates:
                    k2 = re.sub(r"[^a-z0-9]", "", key)
                    if k2 in normalized:
                        return normalized[k2]
                return None
            
            # Find columns using exact same logic as working script
            key_disp = _pick_column(colmap, ("disp #", "disp#", "dispatch #", "dispatch#", "stop id", "stopid", "stop_id", "disp"))
            key_status = _pick_column(colmap, ("status", "pickup status", "state", "pu status"))
            key_asg = _pick_column(colmap, ("asg route", "asgroute", "assigned route", "route", "asg", "asg_route"))
            key_addr_line1 = _pick_column(colmap, (
                "address line 1", "address line1", "address line1.", "address1",
                "addr line 1", "addr1", "address_line_1", "addr_line_1",
                "cust address line 1", "customer address line 1", "street address line 1"
            ))
            key_close = _pick_column(colmap, ("close", "close time", "close_time", "closing time", "customer close", "cust close", "window close", "pickup close", "pu close"))
            key_ready = _pick_column(colmap, ("ready", "ready time", "ready_time", "customer ready", "cust ready", "window open", "open", "pickup ready", "pu ready"))
            key_scan = _pick_column(colmap, ("scan time", "scantime", "scan", "pickup scan", "pu scan", "pud scan", "scanned time", "scan_dt"))
            key_date = _pick_column(colmap, ("p/u date", "pickup date", "pu date", "scheduled date", "sch date", "route date", "date"))
            
            if key_disp is None:
                logging.error(f"Missing required 'Disp #' column. Columns found: {list(df.columns)}")
                return 0, 0, 0, 0, 0
            
            # Use exact same analysis logic as working script
            def _norm(s):
                return s.astype(str).str.strip()
            
            # Active-only mask
            if key_status is not None:
                status_norm = _norm(df[key_status]).str.upper()
                mask_active = status_norm.str.contains("ACTIVE", na=False)
            else:
                mask_active = pd.Series([True] * len(df), index=df.index)
            
            # UNA + Active count
            una_count = 0
            una_tooltip_data = []
            if key_asg is not None:
                mask_una = _norm(df[key_asg]).str.upper().eq("UNA")
                df_una_active = df.loc[mask_una & mask_active]
                una_counts = _norm(df_una_active[key_disp]).replace({"": None}).dropna().value_counts(sort=True)
                una_count = int(una_counts.sum())
                
                # Collect UNA disp#s for tooltip
                una_disp_numbers = _norm(df_una_active[key_disp]).replace({"": None}).dropna().unique()
                una_tooltip_data = sorted(una_disp_numbers)
            
            # Parse times for early/late analysis
            def _parse_time_24(series, date_series=None):
                if series is None:
                    return None
                s = series.astype(str).str.strip()
                
                def fix_hhmm(val):
                    if re.fullmatch(r"\d{3,4}", val or ""):
                        v = val.zfill(4)
                        return f"{v[:2]}:{v[2:]}"
                    return val
                
                s_fixed = s.apply(fix_hhmm)
                if date_series is not None:
                    ds = pd.to_datetime(date_series.astype(str).str.strip(), errors="coerce")
                    combo = ds.dt.strftime("%Y-%m-%d") + " " + s_fixed
                    dt = pd.to_datetime(combo, errors="coerce", format=None)
                else:
                    dummy = "1900-01-01 " + s_fixed
                    dt = pd.to_datetime(dummy, errors="coerce", format=None)
                return dt
            
            close_dt = _parse_time_24(df[key_close], df[key_date] if (key_close and key_date) else None) if key_close else None
            ready_dt = _parse_time_24(df[key_ready], df[key_date] if (key_ready and key_date) else None) if key_ready else None
            scan_dt = _parse_time_24(df[key_scan], df[key_date] if (key_scan and key_date) else None) if key_scan else None
            
            # Status filter for Early/Late: explicitly include CANCELLED and exclude only ACTIVE
            if key_status is not None:
                status_txt = df[key_status].astype(str).str.upper()
                # Explicitly include CANCELLED entries and exclude only ACTIVE
                status_ok = ~status_txt.str.contains("ACTIVE", na=False)
                # Add logging to verify what statuses are being processed
                if status_ok.any():
                    unique_statuses = status_txt[status_ok].unique()
                    logging.info(f"ðŸ” Late counter processing statuses: {unique_statuses}")
            else:
                status_ok = pd.Series([True] * len(df), index=df.index)
            
            # Late pickups count (MUST include CANCELLED stops with scan times)
            late_count = 0
            late_tooltip_data = []
            if (ready_dt is not None) and (close_dt is not None) and (scan_dt is not None):
                late_valid = status_ok & ready_dt.notna() & close_dt.notna() & scan_dt.notna()
                mask_late = late_valid & (scan_dt > close_dt)
                late_count = int(mask_late.sum())
                
                # Log details about what's being counted for late pickups
                if late_count > 0:
                    late_df = df.loc[mask_late]
                    late_statuses = late_df[key_status].astype(str).str.upper().unique() if key_status else ["N/A"]
                    logging.info(f"ðŸ” Late counter found {late_count} entries with statuses: {late_statuses}")
                    # Check specifically for cancelled entries
                    if key_status:
                        cancelled_late = late_df[late_df[key_status].astype(str).str.upper().str.contains("CANCELLED", na=False)]
                        if len(cancelled_late) > 0:
                            logging.info(f"âœ… Late counter includes {len(cancelled_late)} CANCELLED entries with scan times")
                
                # Collect Late pickup data for tooltip (disp# and route)
                if key_asg and late_count > 0:
                    late_df = df.loc[mask_late]
                    for idx in late_df.index:
                        disp = str(late_df.at[idx, key_disp]).strip()
                        route_raw = str(late_df.at[idx, key_asg]).strip()
                        # Format route like test function
                        route_match = re.search(r"(\d+)", route_raw)
                        route = route_match.group(1) if route_match else route_raw
                        late_tooltip_data.append(f"{disp} -> {route}")
            
            # Early pickups count (MUST include CANCELLED stops with scan times)
            early_count = 0
            early_tooltip_data = []
            if (ready_dt is not None) and (scan_dt is not None):
                early_valid = status_ok & ready_dt.notna() & scan_dt.notna()
                mask_early = early_valid & (scan_dt < ready_dt)
                early_count = int(mask_early.sum())
                
                # Collect Early pickup data for tooltip (disp# and route)
                if key_asg and early_count > 0:
                    early_df = df.loc[mask_early]
                    for idx in early_df.index:
                        disp = str(early_df.at[idx, key_disp]).strip()
                        route_raw = str(early_df.at[idx, key_asg]).strip()
                        # Format route like test function
                        route_match = re.search(r"(\d+)", route_raw)
                        route = route_match.group(1) if route_match else route_raw
                        early_tooltip_data.append(f"{disp} -> {route}")
            
            # Duplicate on-call count (simplified version)
            duplicates_count = 0
            duplicates_tooltip_data = {}
            if key_addr_line1 and key_disp:
                # Group by normalized address
                addr_norm = (
                    df[key_addr_line1].fillna("").astype(str).str.lower()
                    .str.replace(r"\s+", " ", regex=True).str.strip()
                )
                work = df.loc[mask_active].copy()
                work["_addr_norm"] = addr_norm.loc[mask_active]
                work["_disp"] = work[key_disp].fillna("").astype(str).str.strip()
                work["_is_oncall"] = work["_disp"].str.startswith("3")
                
                for addr_k, g in work.groupby("_addr_norm", dropna=False):
                    if not addr_k or len(g) < 2:
                        continue
                    oncs_df = g[g["_is_oncall"]]
                    if len(oncs_df) >= 2:  # 2 or more on-calls at same address
                        duplicates_count += len(oncs_df)
                        
                        # Collect duplicate data for tooltip (address and disp#s with routes)
                        if key_asg:
                            addr_duplicates = []
                            for idx in oncs_df.index:
                                disp = str(oncs_df.at[idx, key_disp]).strip()
                                route_raw = str(oncs_df.at[idx, key_asg]).strip()
                                # Format route like test function
                                route_match = re.search(r"(\d+)", route_raw)
                                route = route_match.group(1) if route_match else route_raw
                                addr_duplicates.append(f"{disp} -> Route {route}")
                            duplicates_tooltip_data[addr_k] = addr_duplicates
            
            # FORGE Received blank + Active (excluding UNA) count
            untransmitted_count = 0
            untransmitted_tooltip_data = {}
            if key_asg is not None:
                # Find the FORGE Received column
                key_forge_received = _pick_column(colmap, ("forge received", "forgereceived", "forge_received"))
                
                if key_forge_received is not None:
                    # Count unique routes that are Active, not UNA, and have blank/empty FORGE Received
                    # This matches the test function logic exactly
                    sr = df[key_forge_received]
                    forge_blank = sr.isna() | sr.astype(str).str.strip().eq("")
                    asg_norm = _norm(df[key_asg]).str.upper()
                    mask_untransmitted = forge_blank & mask_active & (~asg_norm.eq("UNA"))
                    
                    # Get unique routes from the filtered data
                    routes_series = df.loc[mask_untransmitted, key_asg].astype(str).str.strip()
                    route_label = routes_series.str.extract(r"(\d{1,})", expand=False)
                    route_label = route_label.where(route_label.notna(), routes_series)
                    route_label = route_label.apply(lambda x: x if re.fullmatch(r"\d+", str(x)) else str(x))
                    
                    # Count unique routes (not total records)
                    untransmitted_routes = route_label.unique()
                    untransmitted_count = len(untransmitted_routes)
                    
                    # Collect route breakdown for tooltip (route and count of untransmitted stops)
                    untransmitted_tooltip_data = route_label.value_counts().to_dict()
                else:
                    # Fallback: count unique routes that are Active but not UNA
                    mask_untransmitted = mask_active & ~_norm(df[key_asg]).str.upper().eq("UNA")
                    routes_series = df.loc[mask_untransmitted, key_asg].astype(str).str.strip()
                    route_label = routes_series.str.extract(r"(\d{1,})", expand=False)
                    route_label = route_label.where(route_label.notna(), routes_series)
                    route_label = route_label.apply(lambda x: x if re.fullmatch(r"\d+", str(x)) else str(x))
                    
                    untransmitted_routes = route_label.unique()
                    untransmitted_count = len(untransmitted_routes)
                    
                    # Collect route breakdown for tooltip (route and count of untransmitted stops)
                    untransmitted_tooltip_data = route_label.value_counts().to_dict()
            
            logging.info(f"ðŸ“Š Analysis complete - UNA: {una_count}, Early: {early_count}, Late: {late_count}, Untransmitted: {untransmitted_count}, Duplicates: {duplicates_count}")
            
            # Return counts and tooltip data
            tooltip_data = {
                'una': una_tooltip_data,
                'early': early_tooltip_data,
                'late': late_tooltip_data,
                'untransmitted': untransmitted_tooltip_data,
                'duplicates': duplicates_tooltip_data
            }
            
            return una_count, early_count, late_count, untransmitted_count, duplicates_count, tooltip_data
            
        except Exception as e:
            logging.error(f"Error analyzing CSV data: {e}")
            import traceback
            logging.error(f"Traceback: {traceback.format_exc()}")
            empty_tooltip_data = {
                'una': [],
                'early': [],
                'late': [],
                'untransmitted': {},
                'duplicates': {}
            }
            return 0, 0, 0, 0, 0, empty_tooltip_data
    
    def _find_column(self, col_map, candidates):
        """Find a column by trying multiple possible names."""
        for candidate in candidates:
            if candidate in col_map:
                return col_map[candidate]
        return None
    
    def _switch_to_default(self, driver):
        """Switch to default content."""
        try:
            driver.switch_to.default_content()
        except:
            pass
    
    def _find_element_across_frames(self, driver, locator):
        """Find element across all frames using the same approach as the working test script."""
        by, value = locator
        self._switch_to_default(driver)
        
        try:
            el = WebDriverWait(driver, 4).until(EC.presence_of_element_located((by, value)))
            return el, None
        except:
            pass
        
        # Search in iframes
        frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
        for idx, _ in enumerate(frames):
            self._switch_to_default(driver)
            try:
                driver.switch_to.frame(idx)
                el = WebDriverWait(driver, 4).until(EC.presence_of_element_located((by, value)))
                return el, idx
            except:
                continue
            finally:
                self._switch_to_default(driver)
        
        return None, None
    
    def _safe_click(self, driver, element):
        """Safe click method from the working test script."""
        try:
            element.click()
            return
        except Exception:
            pass
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'})", element)
            ActionChains(driver).move_to_element(element).pause(0.05).perform()
            element.click()
            return
        except Exception:
            pass
        driver.execute_script("arguments[0].click();", element)
    
    def start_test_scheduler(self):
        """Start the test function scheduler to run every 15 minutes continuously."""
        if self.test_scheduler_active:
            return
            
        self.test_scheduler_active = True
        self.test_scheduler_stop_event.clear()
        
        def scheduler_worker():
            # Wait 15 minutes before first run
            for _ in range(900):
                if self.test_scheduler_stop_event.is_set():
                    break
                time.sleep(1)
                
            while not self.test_scheduler_stop_event.is_set():
                # Show loading spinners before starting test function
                self.after(0, lambda: self._show_loading_counters())
                self.after(0, lambda: self.test_status_label.configure(text="ðŸ”„ Scheduled pickup statistics update..."))
                
                # Run test function in background thread
                threading.Thread(target=self._run_test_function_worker, daemon=True).start()
                
                # Wait 15 minutes (900 seconds) before next run
                for _ in range(900):
                    if self.test_scheduler_stop_event.is_set():
                        break
                    time.sleep(1)
        
        # Start scheduler thread
        threading.Thread(target=scheduler_worker, daemon=True).start()
    
    def stop_test_scheduler(self):
        """Stop the test function scheduler."""
        self.test_scheduler_active = False
        self.test_scheduler_stop_event.set()
        # Status label is hidden, no text needed

    def _on_sof_monitor_ready(self):
        """Hide the SoF+ monitor-startup loading bar."""
        self.sof_monitor_progress.stop()
        self.sof_monitor_progress.pack_forget()    

    def _sof_monitor_worker(self):
        """
        Reuses the single Route Summary panel to check each route's status:
        1. If the "Unclear" button is present â‡’ Grey (#A9A9A9)
        2. Else if last Scan Time â‰¤ 15 min ago â‡’ DarkGreen (#006400)
        3. Else â‡’ Red (#FF0000)
        Shows "" on each button while it's being checked, then restacks.
        """
        import tkinter as tk

        # â”€â”€â”€ Paint helper (must come before any self.after(..., paint, ...) calls) â”€â”€â”€
        def paint(btn, fc, oc, label, cleared):
            try:
                if btn.winfo_exists():
                    cfg = {
                        "fg_color": fc,
                        "state": "normal",  # All routes are clickable
                        "text": label
                    }
                    # only set a border if oc is non-None
                    if oc:
                        cfg["border_color"] = oc
                        # use a 4px border for both red and orange
                        cfg["border_width"] = 4
                    else:
                        cfg["border_width"] = 0
                    btn.configure(**cfg)
                    
                    # Store the route color for use in popups
                    route = int(label) if label.isdigit() else label
                    if route and hasattr(self, 'sof_route_colors'):
                        # Only update if color actually changed
                        if route not in self.sof_route_colors or self.sof_route_colors[route] != fc:
                            self.sof_route_colors[route] = fc
                            


            except tk.TclError:
                pass

        # 1) Start headless Chrome
        driver_path = ChromeDriverManager().install()
        service     = Service(driver_path)
        opts        = Options()
        opts.add_argument("--disable-gpu")
        opts.add_argument("--headless")
        opts.add_argument("--window-size=1920,1080")
        driver      = webdriver.Chrome(service=service, options=opts)
        wait   = WebDriverWait(driver, 20)

        try:
            # 2) FAST home-page detection with MFA via CarrierSeleniumMonitor
            import threading
            login_helper = CarrierSeleniumMonitor(
                username=self.username,
                password=self.password,
                dispatch_location=self.sof_loc,
                stop_event=self.sof_monitor_stop_event,
                message_types=[],  # not used for login
                mfa_factor_callback=self.prompt_mfa_factor,
                mfa_code_callback  =self.prompt_mfa_code
            )
            login_helper.sso_navigate(
                driver, wait,
                "https://internal.example.com", 
                self.username, self.password
            )

            for attempt in range(5):
                for _ in range(20):
                    if "/home" in driver.current_url:
                        break
                    time.sleep(0.1)
                try:
                    WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH,
                            "//span[normalize-space(text())='FRO']"
                        ))
                    )
                    logging.info(f"âœ… SoF+ monitor home detected on attempt {attempt+1}")
                    break
                except TimeoutException:
                    logging.warning(f"â— SoF+ monitor home detection attempt {attempt+1} failed")
                    if attempt < 4:
                        logging.info("ðŸ”„ Refreshing and retryingâ€¦")
                        driver.refresh()
                        time.sleep(1)
                    else:
                        logging.error("âŒ SoF+ monitor: FRO menu never detected; aborting")
                        driver.quit()
                        return
            # â”€â”€â”€ now logged in, proceed to click FRO / Route Summary â”€â”€â”€â”€â”€â”€â”€

            fro_el = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']")))
            driver.execute_script("arguments[0].click();", fro_el)
            rs_el  = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[normalize-space(text())='Route Summary']"
            )))
            driver.execute_script("arguments[0].click();", rs_el)

            # 3) Enter dispatch location once
            loc_input = wait.until(EC.element_to_be_clickable((By.ID, "location")))
            loc_input.clear()
            loc_input.send_keys(self.sof_loc, Keys.RETURN)

            # 4) Grab the route input once
            route_input = wait.until(EC.element_to_be_clickable((By.ID, "route")))

            # â”€â”€â”€ Monitor is now ready â€” hide the startup bar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            self.after(0, self._on_sof_monitor_ready)
            # Update all tooltips with current status information
            self.after(1000, self._update_all_route_tooltips)

            # Wait for page to be ready after location entry, then clear route input
            time.sleep(1)  # Give page time to load after location entry
            route_input.clear()
            time.sleep(0.5)  # Small delay to ensure clear is processed
            
            # Wait for summary table to be present/ready before checking routes
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
                logging.info("âœ… Summary table is ready, starting route checks")
            except TimeoutException:
                logging.warning("âš ï¸ Summary table not found, but continuing with route checks")

            # 5) Monitoring loop
            while not self.sof_monitor_stop_event.is_set():
                # Update the "last update" timestamp at the start of each monitoring cycle
                self._last_sof_update = datetime.now()
                now_min = datetime.now().hour * 60 + datetime.now().minute

                for route, btn in list(self.sof_route_buttons.items()):
                    try:
                        # allow thread to stop early
                        if self.sof_monitor_stop_event.is_set():
                            break

                        # show hourglass while checking
                        self.after(0, lambda b=btn: b.configure(text="â³"))
                        logging.info(f"ðŸ›°ï¸ Checking SoF+ for route {route}")

                        # pad to 3 digits (e.g. '1'â†’'001', '55'â†’'055') before sending
                        padded_route = route.zfill(3)
                        # Clear route input and wait a moment to ensure it's empty
                        route_input.clear()
                        time.sleep(0.2)  # Small delay to ensure clear is processed
                        route_input.send_keys(padded_route, Keys.RETURN)
                        # â”€â”€â”€ Wait for the summary table to load this route â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        padded = route.zfill(3)
                        route_xpath = (
                            f"//td[span[normalize-space(text())='{route}'"
                            f" or normalize-space(text())='{padded}']]"
                        )
                        # â”€â”€â”€ Wait for the summary table to load this route (up to 5Ã—15s) â”€â”€â”€
                        loaded = False
                        for attempt in range(1, 6):
                            try:
                                WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH, route_xpath))
                                )
                                loaded = True
                                break
                            except TimeoutException:
                                logging.warning(
                                    f"[SoF+] attempt {attempt}/5: summary for route {route} did not load in 15s, retryingâ€¦"
                                )
                                # re-enter the route number and try again
                                route_input.clear()
                                route_input.send_keys(padded_route, Keys.RETURN)

                        if not loaded:
                            logging.warning(
                                f"[SoF+] summary for route {route} failed after 5 attempts, skipping."
                            )
                            continue

                        # Give the summary view a brief moment to fully render stops
                        # before we start inspecting buttons/rows, otherwise the very
                        # first route checked can appear to have no stops.
                        time.sleep(1.0)

                        # â”€â”€â”€ Detect route_cleared â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        try:
                            clear_btn = driver.find_element(By.CSS_SELECTOR, "button.clearvisible")
                            route_cleared = (clear_btn.text.strip() == "Unclear")
                        except NoSuchElementException:
                            route_cleared = False

                        # â”€â”€â”€ Detect open stops â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        # Mirror Auto-Clear logic: count only real stop IDs and
                        # ignore the synthetic "LAST STOP" summary row.
                        try:
                            divs = driver.find_elements(By.CSS_SELECTOR, "div.ng-star-inserted")
                            stops_list = []
                            for d in divs:
                                txt = d.text.strip()
                                if not (txt.isdigit() and len(txt) == 4):
                                    continue
                                # Skip any stop that lives in the "LAST STOP" summary row
                                try:
                                    d.find_element(
                                        By.XPATH,
                                        "./ancestor::tr[.//span[@title='Last Stop']]",
                                    )
                                    continue  # this is in the LAST STOP row â†’ ignore
                                except NoSuchElementException:
                                    pass
                                stops_list.append(txt)
                        except Exception:
                            stops_list = []
                        stops_exist = bool(stops_list)

                        # â”€â”€â”€ Detect break_state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        break_state = None
                        try:
                            cand = driver.find_elements(
                                By.XPATH,
                                "//div[contains(normalize-space(.), '(BRK)') "
                                "or contains(normalize-space(.), 'Begin Break') "
                                "or contains(normalize-space(.), 'End Break')]"
                            )
                            for d in cand:
                                txt = d.text.strip().lower()
                                if "(brk)" in txt or "begin break" in txt:
                                    break_state = "begin"
                                    break
                                if "end break" in txt:
                                    break_state = "end"
                                    break
                        except Exception:
                            pass

                        # â”€â”€â”€ Detect "LOCKED" status â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        try:
                            locked = bool(driver.find_elements(
                                By.XPATH,
                                "//span[contains(normalize-space(.), 'LOCKED')]"
                            ))
                        except Exception as e:
                            locked = False
                            logging.warning(f"   [SoF+] locked lookup error: {e!r}")
                        logging.info(f"   [SoF+] locked={locked}")

                        # â”€â”€â”€ Detect "closing soon" (â‰¤30 min) or "past closing time" with stale-safe text read â”€
                        soon = False
                        try:
                            close_hdr = driver.find_element(By.XPATH, "//th[contains(., 'Close')]")
                            close_col = len(close_hdr.find_elements(By.XPATH, "preceding-sibling::th")) + 1
                            for c in driver.find_elements(By.XPATH, f"//table//tr/td[{close_col}]/div"):
                                try:
                                    txt = c.text.strip()
                                except StaleElementReferenceException:
                                    continue
                                m = re.match(r"(\d{1,2}):(\d{2})", txt)
                                if m:
                                    close_min = int(m.group(1)) * 60 + int(m.group(2))
                                    # Check if closing within 30 minutes OR if past closing time
                                    if (close_min - now_min) <= 30:
                                        soon = True
                                        break
                        except NoSuchElementException:
                            pass
                        logging.info(f"   [SoF+] soon={soon}")

                        # â”€â”€â”€ Decide fill color & log change â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        if stops_exist and not locked:
                            new_color = "#FF0000"
                            reason    = "stops exist but route not locked"
                        elif route_cleared:
                            new_color = "#006400"
                            reason    = "cleared"
                        elif break_state == "begin":
                            new_color = "#CCCC00"
                            reason    = "begin break (dark yellow)"
                        else:
                            new_color = "#404040"
                            reason    = "default/dark grey"

                        logging.info(
                            f"   [SoF+] route {route}: "
                            f"route_cleared={route_cleared}, "
                            f"break_state={break_state}, "
                            f"stops_exist={stops_exist}, "
                            f"locked={locked}, "
                            f"chosen_color={new_color} ({reason})"
                        )

                        # â”€â”€â”€ Store complete route state information â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        route_state = {
                            'locked': locked,
                            'stops_exist': stops_exist,
                            'break_state': break_state,
                            'route_cleared': route_cleared,
                            'soon': soon,
                            'chosen_color': new_color
                        }
                        
                        # Update the route states dictionary
                        if not hasattr(self, 'sof_route_states'):
                            self.sof_route_states = {}
                        
                        # Check if route status changed (became a problem or stopped being a problem)
                        old_state = self.sof_route_states.get(route, {})
                        old_color = old_state.get('chosen_color', '')
                        is_new_problem = (new_color == '#FF0000' and old_color != '#FF0000')  # Became critical
                        is_new_warning = (new_color == '#CCCC00' and old_color != '#CCCC00')  # Became warning
                        is_no_longer_problem = ((old_color == '#FF0000' or old_color == '#CCCC00') and 
                                               new_color not in ['#FF0000', '#CCCC00'])  # No longer problem
                        
                        self.sof_route_states[route] = route_state
                        
                        # If route status changed and widget is visible, update ticker dynamically
                        if hasattr(self, 'widget_panel_visible') and self.widget_panel_visible:
                            if is_new_problem:
                                self.after(0, self.add_route_to_ticker, route, "critical")
                            elif is_new_warning:
                                self.after(0, self.add_route_to_ticker, route, "warning")
                            elif is_no_longer_problem:
                                self.after(0, self.remove_route_from_ticker, route)
                        
                        # â”€â”€â”€ Paint with outline logic (unchanged) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                        self.after(
                            0,
                            paint,
                            btn,
                            new_color,
                            "#FF0000" if soon else (BRAND_ORANGE if stops_exist else None),
                            route,
                            route_cleared
                        )
                        
                        # Update tooltip with current status
                        self.after(0, self._update_route_tooltip, str(route))
                        
                        # Update the "last update" timestamp after each route to prevent watchdog from thinking we're stale
                        self._last_sof_update = datetime.now()

                    except Exception as e:
                        logging.exception(f"âš ï¸ SoF+ per-route error for {route}, skipping: {e!r}")
                        continue

                # restack & sleepâ€¦
                # â”€â”€â”€ record that SoF+ is still alive â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                self.after(0, self.restack_sof_buttons)
                # Note: _last_sof_update is now updated after each route, so we don't need to set it here
                for _ in range(900):  # 15 min pause (back to normal)
                    if self.sof_monitor_stop_event.is_set():
                        break
                    time.sleep(1)
                
                # Continue to next iteration to check routes again
                continue

        finally:
            driver.quit()
            # Reset the running flag when the worker thread finishes
            if hasattr(self, "sof_monitor_running"):
                self.sof_monitor_running = False
            logging.info("ðŸ”„ SOF+ monitor worker thread finished")    




    def restack_sof_buttons(self):
        """
        Clears and re-grids route buttons with Problem Routes at the top,
        then normal routes sorted numerically with group headers.
        """
        cols = 8
        # clear out whatever's there now
        for w in self.sof_grid_frame.winfo_children():
            w.grid_forget()

        # Separate problem routes from normal routes
        problem_routes = []
        normal_routes = []
        
        for route, btn in self.sof_route_buttons.items():
            try:
                # Check if button has red border or red fill
                fg_color = btn.cget("fg_color")
                border_color = btn.cget("border_color")
                border_width = btn.cget("border_width")
                
                # Red fill (#FF0000) or red border with width > 0
                is_problem = (fg_color == "#FF0000" or 
                            (border_color == "#FF0000" and border_width > 0))
                
                if is_problem:
                    problem_routes.append(route)
                else:
                    normal_routes.append(route)
            except Exception:
                # If we can't determine the button state, treat as normal
                normal_routes.append(route)

        # Sort both lists numerically
        problem_routes.sort(key=lambda r: int(r))
        normal_routes.sort(key=lambda r: int(r))

        row = 0
        col_idx = 0
        half_btn = 20  # half of your button height (40)

        # First, place Problem Routes section if there are any
        if problem_routes:
            # Problem Routes header
            problem_header = ctk.CTkLabel(
                self.sof_grid_frame,
                text="ðŸš¨ PROBLEM ROUTES ðŸš¨",
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color="#FF0000"
            )
            problem_header.grid(row=row, column=0, columnspan=cols, pady=(half_btn, 4))
            row += 1

            # Place problem route buttons
            for route in problem_routes:
                btn = self.sof_route_buttons[route]
                btn.grid(row=row, column=col_idx, padx=5, pady=5)
                
                col_idx += 1
                if col_idx >= cols:
                    col_idx = 0
                    row += 1

            # Add spacing after problem routes section
            if col_idx > 0:  # If we didn't end on a full row, move to next row
                row += 1
            row += 1  # Extra spacing
            col_idx = 0

        # Then place normal routes with group headers
        current_group = None
        for route in normal_routes:
            group = int(route) // 100

            # when you hit a new 100s bucket, insert spacing + header
            if group != current_group:
                # add blank row for spacing (acts as half-button gap)
                if current_group is not None:
                    row += 1

                # group header
                label_text = f"{group*100}-{group*100+99}"
                sep = ctk.CTkLabel(
                    self.sof_grid_frame,
                    text=label_text,
                    font=ctk.CTkFont(size=14, weight="bold"),
                    text_color=BRAND_ORANGE
                )
                # 20px top padding before header, 4px below
                sep.grid(row=row, column=0, columnspan=cols, pady=(half_btn, 4))

                row += 1
                col_idx = 0
                current_group = group

            # place the route button
            btn = self.sof_route_buttons[route]
            btn.grid(row=row, column=col_idx, padx=5, pady=5)

            col_idx += 1
            # wrap to next row every `cols`
            if col_idx >= cols:
                col_idx = 0
                row += 1

    def show_access_denied(self):
        # Create a borderless top-level
        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.grab_set()

        # Outer border
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        # Inner content frame
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        # Header with title and close button
        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Access Denied",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(side="left", padx=10, pady=5)

        def close_popup():
            if messagebox.askyesno("Confirm", "Are you sure you want to close?"):
                popup.destroy()
        ctk.CTkButton(
            header, text="âœ•", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close_popup
        ).pack(side="right", padx=10, pady=5)

        # Allow dragging of the popup
        def start_move(event):
            popup.x0 = event.x; popup.y0 = event.y
        def do_move(event):
            x = popup.winfo_x() + (event.x - popup.x0)
            y = popup.winfo_y() + (event.y - popup.y0)
            popup.geometry(f"+{x}+{y}")
        header.bind("<ButtonPress-1>", start_move)
        header.bind("<B1-Motion>", do_move)

        # Message content
        content = ctk.CTkFrame(inner, fg_color="black")
        content.pack(fill="both", expand=True, padx=10, pady=(0,10))
        ctk.CTkLabel(
            content,
            text="You are not in Dispatch / Southern Region.\n\n"
                 "If you believe this is an error, please contact your manager.",
            font=ctk.CTkFont(size=12),
            text_color="white",
            wraplength=300,
            justify="left"
        ).pack(expand=True, fill="both", pady=10)

        # Center the popup over the main window
        self.update_idletasks()
        mx, my = self.winfo_x(), self.winfo_y()
        mw, mh = self.winfo_width(), self.winfo_height()
        popup.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        x = mx + (mw - pw)//2; y = my + (mh - ph)//2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")    

    def cleanup_processes(self):
        parent = psutil.Process(os.getpid())
        for child in parent.children(recursive=True):
            try:
                child.kill()
            except:
                pass
    
    def start_audit_monitor(self, loc):
        # if there's already a monitor OR nothing to audit, do nothing
        if loc in self.running_audit_monitors or not self.clear_audit_routes.get(loc):
            return

        stop_evt = threading.Event()

        def audit_worker():
            # 0) Log start and routes
            logging.info(f"ðŸš§ Audit monitor START for {loc}: {self.clear_audit_routes[loc]}")
            # --- ensure we can call the same _get_stops helper as Auto-Clear ---
            mon = self.running_clear_monitors.get(loc, {}).get("monitor")
            # 1) Prepare & login with up to 3 retries (mirrors your SoF+ worker)
            # â”€â”€ install chromedriver & silence its logs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            driver_path = ChromeDriverManager().install()
            service     = Service(driver_path, log_path=os.devnull)

            # â”€â”€ build headless-chrome options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            opts = Options()
            opts.add_argument("--disable-extensions")
            opts.add_argument("--headless")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--disable-gpu")

            # â”€â”€ QUIET CHROME FLAGS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            opts.add_argument("--disable-background-networking")
            opts.add_argument("--disable-gcm")
            opts.add_argument("--log-level=3")
            opts.add_experimental_option("excludeSwitches", ["enable-logging"])

            driver = webdriver.Chrome(service=service, options=opts)
            wait   = WebDriverWait(driver, 20)
        
            # â”€â”€â”€ define the URL for SSO login â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            login_url = "https://internal.example.com"

            # â”€â”€â”€ perform the SSO/Identity Portal login (with MFA) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            login_helper = CarrierSeleniumMonitor(
                username=self.username,
                password=self.password,
                dispatch_location=loc,
                stop_event=stop_evt,
                message_types=[],
                mfa_factor_callback=self.prompt_mfa_factor,
                mfa_code_callback  =self.prompt_mfa_code
            )
            try:
                login_helper.sso_navigate(driver, wait, login_url, self.username, self.password)
            except TimeoutException:
                logging.error("âŒ Audit login failed; stopping audit monitor")
                driver.quit()
                self.running_audit_monitors.pop(loc, None)
                return
        
            # 2) Now navigate FRO â†’ Route Summary
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH,
                                                   "//span[normalize-space(text())='Route Summary']"))).click()

            # 2) monitor loop (every 5 minutes)
            while not stop_evt.is_set() and self.clear_audit_routes.get(loc):
                logging.info(f"ðŸ”„ Audit check for {loc}, routes: {self.clear_audit_routes[loc]}")
                for route in list(self.clear_audit_routes[loc]):
                    # fill location
                    loc_field = wait.until(EC.element_to_be_clickable((By.ID, "location")))
                    loc_field.clear()
                    loc_field.send_keys(loc + Keys.RETURN)
                    # fill route
                    route_field = wait.until(EC.element_to_be_clickable((By.ID, "route")))
                    route_field.clear()
                    route_field.send_keys(route + Keys.RETURN)
                    logging.info(f"   Submitted route {route}, waiting for summaryâ€¦")
        
                    # â”€â”€â”€ exactly like Auto-Clear: wait for the summary table to show this route â”€â”€â”€
                    padded = route.zfill(4)
                    xpath_route_td = (
                        f"//td[span[normalize-space(text())='{route}'"
                        f" or normalize-space(text())='{padded}']]"
                    )
                    try:
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.XPATH, xpath_route_td))
                        )
                        logging.info(f"âœ… Found route {route} in summary table")
                    except TimeoutException:
                        logging.warning(f"ðŸš§ Route {route} not found in summary; skipping stop-check")
                        continue

                    # Give the page time to fully render stops (same as Auto-Clear)
                    time.sleep(3)

                    # Click the route row to expand it (stops may be in expandable section)
                    try:
                        route_td = driver.find_element(By.XPATH, xpath_route_td)
                        row = route_td.find_element(By.XPATH, "./ancestor::tr")
                        row.click()
                        time.sleep(2)
                    except Exception:
                        pass

                    # Now detect stops - try multiple selectors (stops can be in div, td, span)
                    def _get_stops_local(drv, route_num):
                        """Detect 4-digit stop IDs from the route summary page.
                        Ignores the synthetic 'LAST STOP' summary row (same as _get_stops).
                        """
                        route_padded = route_num.zfill(4)
                        seen = set()
                        for attempt in range(3):
                            try:
                                for selector in ["div.ng-star-inserted", "td", "span"]:
                                    els = drv.find_elements(By.CSS_SELECTOR, selector)
                                    for el in els:
                                        txt = el.text.strip()
                                        if (txt.isdigit() and len(txt) == 4 and txt != route_padded
                                                and txt not in seen):
                                            # Skip any stop that lives in the "LAST STOP" summary row
                                            try:
                                                el.find_element(
                                                    By.XPATH,
                                                    "./ancestor::tr[.//span[@title='Last Stop']]",
                                                )
                                                continue  # in LAST STOP row â†’ ignore
                                            except NoSuchElementException:
                                                pass
                                            seen.add(txt)
                                if seen:
                                    return list(seen)
                            except StaleElementReferenceException:
                                time.sleep(1)
                            time.sleep(1)
                        return list(seen)

                    stops = _get_stops_local(driver, route)
                    logging.info(f"   Audit sees stops for {route}: {stops}")
                    if not stops:
                        logging.info(f"ðŸš® Audit: clearing route {route}")
                        # record that an audited route was auto-cleared
                        self.log_action(f"{loc}: Audited route {route} cleared")

                        clear_ok = False
                        try:
                            # â”€â”€â”€ No open stops â€“ clearing or blocking route â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            logging.info("ðŸš® No open stops â€“ entering clear/block flow")

                            # 1) Find the Clear/Unclear button
                            try:
                                clear_btn = wait.until(EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, "button.clearvisible"
                                )))
                            except TimeoutException:
                                logging.error("âŒ Clear/Unclear button not found; skipping clear flow")
                                continue

                            btn_text = clear_btn.text.strip()
                            if btn_text == "Clear":
                                # â”€â”€â”€ Click Clear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                for _ in range(3):
                                    try:
                                        clear_btn.click()
                                        break
                                    except StaleElementReferenceException:
                                        time.sleep(0.5)

                                # â”€â”€â”€ Fill clearTime if empty â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                try:
                                    clear_time_input = wait.until(EC.visibility_of_element_located((
                                        By.CSS_SELECTOR, "input[formcontrolname='clearTime']"
                                    )))
                                    if not clear_time_input.get_attribute("value").strip():
                                        clear_time_input.clear()
                                        clear_time_input.send_keys(datetime.now().strftime("%H%M"))
                                except TimeoutException:
                                    logging.warning("âš ï¸ Clear-time input not found; skipping time fill")

                                # â”€â”€â”€ Confirm Clear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                                ok_clear = wait.until(EC.element_to_be_clickable((
                                    By.XPATH,
                                    "//button[contains(@class,'pickupButtons') and normalize-space(text())='OK']"
                                )))
                                ok_clear.click()

                            elif btn_text == "Unclear":
                                logging.info(f"â„¹ï¸ Button text is 'Unclear' for route {route}; skipping Clear click")
                            else:
                                logging.warning(f"âš ï¸ Unexpected button text {btn_text!r}; attempting Clear anyway")
                                try:
                                    clear_btn.click()
                                except Exception as e:
                                    logging.warning(f"âš ï¸ Clear click failed: {e!r}")

                            # 2) Click Block â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            try:
                                block_btn = wait.until(EC.element_to_be_clickable((
                                    By.CSS_SELECTOR, "button.blockvisible"
                                )))
                                block_btn.click()
                            except TimeoutException:
                                logging.warning("âš ï¸ Block button not found; skipping block step")

                            # 3) Fill blockTime if empty â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            try:
                                block_time_input = wait.until(EC.visibility_of_element_located((
                                    By.CSS_SELECTOR, "input[formcontrolname='blockTime']"
                                )))
                                if not block_time_input.get_attribute("value").strip():
                                    block_time_input.clear()
                                    block_time_input.send_keys(datetime.now().strftime("%H%M"))
                            except TimeoutException:
                                logging.warning("âš ï¸ Block-time input not found; skipping time fill")

                            # 4) Confirm Block â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            for _ in range(3):
                                try:
                                    ok_block = wait.until(EC.element_to_be_clickable((
                                        By.XPATH,
                                        "//button[contains(@class,'pickupButtons') and normalize-space(text())='OK']"
                                    )))
                                    ok_block.click()
                                    break
                                except StaleElementReferenceException:
                                    time.sleep(0.5)

                            # 5) Mark as cleared, log â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                            if mon:
                                mon.cleared_routes.add(route)
                                ts = datetime.now()
                                mon.history.append((route, ts))
                                logging.info(f"âœ… Route {route} cleared at {ts:%H:%M:%S}")
                                if mon.log_callback:
                                    mon.log_callback(f"{mon.dispatch_location}: Route {route} cleared at {ts:%H:%M:%S}")
                                if mon.update_status:
                                    mon.update_status(mon.dispatch_location, f"Cleared {route} at {ts:%H:%M:%S}")
                            else:
                                ts = datetime.now()
                                logging.info(f"âœ… Route {route} cleared at {ts:%H:%M:%S}")
                                logging.warning("âš ï¸ No monitor for location; clear reply not sent")

                            clear_ok = True
                        except Exception as e:
                            logging.warning(f"âš ï¸ Audit clear failed for {route}: {e!r}")

                        # After a successful clear/block flow, attempt to send the clear reply
                        # using the *Autoâ€‘Clear* monitor's browser via CREATE MESSAGE so we
                        # don't depend on the original REQUEST TO CLEAR row still existing.
                        if clear_ok and mon and getattr(mon, "driver", None):
                            try:
                                logging.info(
                                    f"âœ‰ï¸ Attempting to send audit clear reply for route {route} via Autoâ€‘Clear driver (CREATE MESSAGE)"
                                )
                                mon._send_clear_via_create_message(
                                    mon.driver,
                                    route,
                                    mon._format_clear_reply(),
                                )
                                logging.info(f"âœ‰ï¸ Audit clear reply sent for route {route} via CREATE MESSAGE")
                            except Exception as e:
                                logging.warning(
                                    f"âš ï¸ Audit clear for {route} succeeded, but clear reply could not be sent via Autoâ€‘Clear driver (CREATE MESSAGE): {e!r}"
                                )
                
                        # 4) Return to Messages tab and Refresh filter
                        for sel in (INBOX_TAB_XPATH, MSG_TAB_XPATH, REFRESH_BTN_CSS):
                            try:
                                if sel.startswith("span"):
                                    driver.find_element(By.CSS_SELECTOR, sel).click()
                                else:
                                    driver.find_element(By.XPATH, sel).click()
                            except:
                                pass
                        time.sleep(1)
                
                        # 5) Stop auditing this route
                        logging.info(f"ðŸ—‘ï¸ Removing {route} from audit list")
                        self.clear_audit_routes[loc].remove(route)
                    else:
                        # still open stops â†’ show warning popup
                        self.after(0, lambda r=route, s=stops: 
                            self.show_audit_warning(loc, r, s)
                        )
                # if still have routes, wait 5 minutes; else fall out
                if self.clear_audit_routes.get(loc):
                    stop_evt.wait(300)
            driver.quit()
            self.running_audit_monitors.pop(loc, None)
            logging.info(f"ðŸ›‘ Audit monitor for {loc} stopped")

        # 3) launch it
        t = threading.Thread(target=audit_worker, daemon=True)
        self.running_audit_monitors[loc] = {"thread": t, "stop_event": stop_evt}
        t.start()
        # record an initial "last update" so watchdog will notice if run() hangs
        self._last_clear_update[loc] = datetime.now()

    def stop_audit_monitor(self, loc):
        info = self.running_audit_monitors.pop(loc, None)
        if info:
            info["stop_event"].set()

    def show_audit_window(self, loc: str):
        """Show audit status window for a location."""
        # create a true top-level, borderless, always-on-top dialog
        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.grab_set()

        # outer border frame
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        # header with title and close button
        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text=f"Audit Routes for {loc}",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(side="left", padx=10, pady=5)
        def close_popup():
            popup.grab_release()
            popup.destroy()
        ctk.CTkButton(
            header, text="âœ•", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close_popup
        ).pack(side="right", padx=10, pady=5)
        # allow dragging by header
        def start_move(evt):
            popup._drag_x = evt.x; popup._drag_y = evt.y
        def do_move(evt):
            x = evt.x_root - popup._drag_x
            y = evt.y_root - popup._drag_y
            popup.geometry(f"+{x}+{y}")
        header.bind("<ButtonPress-1>", start_move)
        header.bind("<B1-Motion>", do_move)

        # entry row for adding a route
        entry_frame = ctk.CTkFrame(inner, fg_color="black")
        entry_frame.pack(fill="x", pady=(10, 0), padx=10)
        entry = ctk.CTkEntry(entry_frame, placeholder_text="Route #", width=100)
        entry.pack(side="left", padx=(0,10))
        entry.bind("<Return>", lambda e: add_route())
        def add_route():
            val = entry.get().strip()
            if val.isdigit():
                num = int(val)
                self.clear_audit_routes.setdefault(loc, set()).add(str(num))
                self.start_audit_monitor(loc)
                self.log_action(f"{loc}: route {num} added to audit list")
                entry.delete(0, "end")
                populate_list()
            else:
                messagebox.showerror("Invalid Route", "Please enter a valid route number.")
        ctk.CTkButton(
            entry_frame, text="Add",
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=add_route
        ).pack(side="left")

        # scrollable list of audit routes
        list_frame = ctk.CTkScrollableFrame(inner, fg_color="black", height=150)
        list_frame.pack(fill="both", expand=False, padx=10, pady=(10,10))
        def populate_list():
            # clear out old entries
            for w in list_frame.winfo_children():
                w.destroy()
            for num in sorted(self.clear_audit_routes.get(loc, []), key=int):
                row = ctk.CTkFrame(list_frame, fg_color="black")
                row.pack(fill="x", pady=2)
                row.grid_columnconfigure(0, weight=1)
                # centered route number
                ctk.CTkLabel(
                    row,
                    text=str(num),
                    font=ctk.CTkFont(size=12),
                    text_color="white"
                ).grid(row=0, column=0, sticky="ew")
                # remove button
                def _remove(r=num, rf=row):
                    if not messagebox.askyesno(
                        "Confirm Removal",
                        f"Remove route {r} from audit list?"
                    ):
                        return
                    self.clear_audit_routes[loc].discard(str(r))
                    if not self.clear_audit_routes[loc]:
                        self.stop_audit_monitor(loc)
                    self.log_action(f"{loc}: route {r} removed from audit list")
                    rf.destroy()
                ctk.CTkButton(
                    row,
                    text="âœ•",
                    width=20, height=20,
                    fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                    command=_remove
                ).grid(row=0, column=1, padx=5)
        populate_list()

        # center the popup over the main window
        self.update_idletasks()
        mx, my = self.winfo_rootx(), self.winfo_rooty()
        mw, mh = self.winfo_width(), self.winfo_height()
        popup.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        x = mx + (mw - pw) // 2
        y = my + (mh - ph) // 2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")
        popup.lift()
        popup.focus_force()

    def show_audit_warning(self, loc: str, route: str, stops: list):
        """Styled, always-on-top pop-up matching other CTkToplevel dialogs."""
        key = f"{loc}_{route}"
        # If a popup for this route already exists, just bring it forward
        existing = self._audit_popups.get(key)
        if existing and existing.winfo_exists():
            existing.lift()
            existing.focus_force()
            return

        # Otherwise create a new one and remember it
        popup = ctk.CTkToplevel(self)
        self._audit_popups[key] = popup
        popup.overrideredirect(True)
        popup.grab_set()
        popup.attributes("-topmost", True)

        # Outer border + inner frame
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        # Header with title + close button + drag handling
        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Audit Warning",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(side="left", padx=10, pady=5)

        def close():
            popup.grab_release()
            popup.destroy()
            self._audit_popups.pop(key, None)

        ctk.CTkButton(
            header, text="âœ•", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close
        ).pack(side="right", padx=10, pady=5)

        # enable dragging by header
        def start_move(evt):
            popup._drag_x = evt.x; popup._drag_y = evt.y
        def do_move(evt):
            nx = evt.x_root - popup._drag_x
            ny = evt.y_root - popup._drag_y
            popup.geometry(f"+{nx}+{ny}")
        header.bind("<ButtonPress-1>", start_move)
        header.bind("<B1-Motion>",    do_move)

        # Content area
        content = ctk.CTkFrame(inner, fg_color="black")
        content.pack(fill="both", expand=True, padx=10, pady=(0,10))
        msg = f"Route {route} in {loc} still has open stops:\n{stops}"
        ctk.CTkLabel(
            content,
            text=msg,
            font=ctk.CTkFont(size=12),
            text_color="white",
            wraplength=380,
            justify="left"
        ).pack(fill="both", expand=True, pady=(5,10))

        # OK button
        ctk.CTkButton(
            inner, text="OK",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close
        ).pack(pady=(0, 10))

        # â”€â”€â”€ Center the popup over the main window â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.update_idletasks()
        # Main window's top-left corner
        mx, my = self.winfo_rootx(), self.winfo_rooty()
        # Main window's size
        mw, mh = self.winfo_width(), self.winfo_height()
        # Popup's requested size
        popup.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        # Compute centered position
        x = mx + (mw - pw) // 2
        y = my + (mh - ph) // 2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")
        popup.lift()
        popup.focus_force()

    def _check_watchdog(self):
        now = datetime.now()
        stale_delta = timedelta(minutes=2)
        stale_delta_sof = timedelta(minutes=15)

        # â€”â€” Auto-Read monitors â€”â€” 
        for loc, info in list(self.running_monitors.items()):
            last = self._last_read_update.get(loc)
            # If it's been >5 min with no updatesâ€”or thread diedâ€”restart it
            if (not info["thread"].is_alive()) or (last and now - last > stale_delta):
                logging.warning(f"âš ï¸ Auto-Read for {loc} staleâ€”restarting")
                frame = info["status_label"].master
                # grab the same settings you originally used:
                types      = info["monitor"].message_types
                reply_mode = info["monitor"].reply_mode
                # tear down
                self.stop_monitoring(loc, frame)
                # re-launch
                # (re-insert into the dispatch box so add_location picks up)
                self.dispatch_entry.insert(0, loc)
                # restore the checkboxes & switch:
                self.pickup_reminder_var .set("Pickup Reminder" in types)
                self.early_pu_var        .set("Early PU" in types)
                self.no_pickup_list_var  .set("No Pickup List" in types)
                self.reply_on_pr_var     .set(reply_mode)
                self.add_location(False)

        # â€”â€” Auto-Clear monitors â€”â€” 
        for loc, info in list(self.running_clear_monitors.items()):
            last = self._last_clear_update.get(loc)
            if (not info["thread"].is_alive()) or (last and now - last > stale_delta):
                logging.warning(f"âš ï¸ Auto-Clear for {loc} staleâ€”restarting")
                frame = info["frame"]
                reply = info["monitor"].clear_reply
                # tear down
                self.stop_clear_monitor(loc, frame)
                # re-launch
                self.clear_dispatch_entry.insert(0, loc)
                self.clear_reply_entry.delete(0, "end")
                self.clear_reply_entry.insert(0, reply)
                self.add_clear_location(False)

        # â€”â€” SoF+ monitor â€”â€” 
        if hasattr(self, "sof_monitor_stop_event") and not self.sof_monitor_stop_event.is_set():
            last = self._last_sof_update
            if last and now - last > stale_delta_sof:
                logging.warning("âš ï¸ SoF+ monitor staleâ€”restarting monitor thread only")
                # 1) Signal the old monitor to quit
                self.sof_monitor_stop_event.set()
                # 2) (Optional) give it a moment to shut down
                time.sleep(1)
                # 3) Kick off a fresh monitor on the same loaded routes
                self.start_sof_monitor()

        # reschedule next check
        self.after(60_000, self._check_watchdog)

    def handle_login_failure(self, error_msg="Incorrect username or password.\nPlease try again."):
        # Show a popup if we have an error message
        messagebox.showerror("Login Failed", error_msg)
        # 1) Stop & remove all Auto-Read monitors
        for loc, info in list(self.running_monitors.items()):
            frame = info["status_label"].master
            self.stop_monitoring(loc, frame)

        # 2) Stop & remove all Auto-Clear monitors
        for loc, info in list(self.running_clear_monitors.items()):
            frame = info["frame"]
            self.stop_clear_monitor(loc, frame)

        # 3) Destroy the dispatch UI
        if self.dispatch_frame:
            self.dispatch_frame.destroy()
            self.dispatch_frame = None
            # **also reset these so init_dispatch_ui() will rebuild them**
            self.locations_container = None
            self.clear_locations_container = None

        # 4) Show the login screen again
        self.init_login_ui()

    def log_action(self, message: str):
        """Append a timestamped entry to the Action Log tab."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}\n"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", entry)
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def show_skip_popup(self, loc):
        # create a true top-level, borderless, always-on-top dialog
        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.grab_set()

        # outer border frame
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        # header with title and close button
        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text=f"Skip Clear Routes for {loc}",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(side="left", padx=10, pady=5)
        def close_popup():
            popup.grab_release()
            popup.destroy()
        ctk.CTkButton(
            header, text="âœ•", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close_popup
        ).pack(side="right", padx=10, pady=5)
        # allow dragging by header
        def start_move(evt):
            popup._drag_x = evt.x; popup._drag_y = evt.y
        def do_move(evt):
            x = evt.x_root - popup._drag_x
            y = evt.y_root - popup._drag_y
            popup.geometry(f"+{x}+{y}")
        header.bind("<ButtonPress-1>", start_move)
        header.bind("<B1-Motion>", do_move)

        # entry row for adding a route
        entry_frame = ctk.CTkFrame(inner, fg_color="black")
        entry_frame.pack(fill="x", pady=(10, 0), padx=10)
        entry = ctk.CTkEntry(entry_frame, placeholder_text="Route #", width=100)
        entry.pack(side="left", padx=(0,10))
        entry.bind("<Return>", lambda e: add_route())
        def add_route():
            val = entry.get().strip()
            if val.isdigit():
                num = int(val)
                self.skip_clear_routes.setdefault(loc, set()).add(num)
                self.log_action(f"{loc}: route {num} added to skip list")
                entry.delete(0, "end")
                populate_list()
        ctk.CTkButton(
            entry_frame, text="Add",
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=add_route
        ).pack(side="left")

        # scrollable list of skip routes
        list_frame = ctk.CTkScrollableFrame(inner, fg_color="black", height=150)
        list_frame.pack(fill="both", expand=False, padx=10, pady=(10,10))
        def populate_list():
            # clear out old entries
            for w in list_frame.winfo_children():
                w.destroy()
            for num in sorted(self.skip_clear_routes.get(loc, [])):
                row = ctk.CTkFrame(list_frame, fg_color="black")
                row.pack(fill="x", pady=2)
                row.grid_columnconfigure(0, weight=1)
                # centered route number
                ctk.CTkLabel(
                    row,
                    text=str(num),
                    font=ctk.CTkFont(size=12),
                    text_color="white"
                ).grid(row=0, column=0, sticky="ew")
                # remove button
                def _remove(r=num, rf=row):
                    if not messagebox.askyesno(
                        "Confirm Removal",
                        f"Remove route {r} from skip list?"
                    ):
                        return
                    self.skip_clear_routes[loc].remove(r)
                    self.log_action(f"{loc}: route {r} removed from skip list")
                    rf.destroy()
                ctk.CTkButton(
                    row,
                    text="âœ•",
                    width=20, height=20,
                    fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
                    command=_remove
                ).grid(row=0, column=1, padx=5)
        populate_list()

        # center the popup over the main window
        self.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        mx, my = self.winfo_x(), self.winfo_y()
        mw, mh = self.winfo_width(), self.winfo_height()
        x = mx + (mw - pw)//2; y = my + (mh - ph)//2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")

    def prompt_mfa_factor(self) -> str:
        choice = {"value": None}

        # 1) create hidden popup
        popup = ctk.CTkToplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)

        # --- build header ---
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="SSO",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(pady=(12,0))
        ctk.CTkLabel(
            header,
            text="Verification",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(pady=(0,12))

        def close():
            popup.destroy()
        close_btn = ctk.CTkButton(
            header, text="âœ•", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close
        )
        close_btn.place(relx=1.0, x=-10, y=10, anchor="ne")

        # dragging
        header.bind("<ButtonPress-1>", lambda e: setattr(popup, "_drag_x", e.x) or setattr(popup, "_drag_y", e.y))
        header.bind("<B1-Motion>",    lambda e: popup.geometry(f"+{popup.winfo_x() + (e.x - popup._drag_x)}+{popup.winfo_y() + (e.y - popup._drag_y)}"))

        # --- build buttons ---
        content = ctk.CTkFrame(inner, fg_color="black")
        content.pack(fill="both", expand=True, padx=20, pady=(0,20))

        def pick(factor):
            choice["value"] = factor
            popup.destroy()

        btn_font = ctk.CTkFont(size=14, weight="bold")

        enter_btn = ctk.CTkButton(
            content,
            text="Enter Code",
            font=btn_font,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=lambda: pick("code")
        )
        enter_btn.pack(fill="x", pady=(0,10))

        push_btn = ctk.CTkButton(
            content,
            text="Push Notification",
            font=btn_font,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=lambda: pick("push")
        )
        push_btn.pack(fill="x")

        # --- size & center behind main window ---
        self.update_idletasks()
        pw, ph = popup.winfo_reqwidth(), popup.winfo_reqheight()
        mx, my = self.winfo_rootx(), self.winfo_rooty()
        mw, mh = self.winfo_width(), self.winfo_height()
        x = mx + (mw - pw)//2
        y = my + (mh - ph)//2
        popup.geometry(f"{pw}x{ph}+{x}+{y}")

        # now show it behind, then raise & grab
        popup.deiconify()
        popup.lower(self)
        popup.attributes("-topmost", True)
        popup.grab_set()

        popup.wait_window()
        return choice["value"]

    # â”€â”€â”€ New MFA Code-Entry Pop-up â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def prompt_mfa_code(self) -> str:
        code_var = tk.StringVar()

        # 1) create hidden popup
        popup = ctk.CTkToplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self)

        # --- build UI ---
        outer = ctk.CTkFrame(popup, fg_color="#404040", corner_radius=0)
        outer.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(outer, fg_color="black", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        header = ctk.CTkFrame(inner, fg_color="black")
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="Enter",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(pady=(12,0))
        ctk.CTkLabel(
            header,
            text="SSO Code",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(pady=(0,12))

        def close():
            popup.destroy()

        close_btn = ctk.CTkButton(
            header, text="âœ•", width=30, height=25,
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close
        )
        close_btn.place(relx=1.0, x=-10, y=10, anchor="ne")

        header.bind("<ButtonPress-1>", lambda e: setattr(popup, "_drag_x", e.x) or setattr(popup, "_drag_y", e.y))
        header.bind("<B1-Motion>",    lambda e: popup.geometry(f"+{popup.winfo_x() + (e.x - popup._drag_x)}+{popup.winfo_y() + (e.y - popup._drag_y)}"))

        content = ctk.CTkFrame(inner, fg_color="black")
        content.pack(fill="both", expand=True, padx=20, pady=(0,20))
        entry = ctk.CTkEntry(content, textvariable=code_var, placeholder_text="123456")
        entry.pack(pady=(0,10))
        entry.focus()
        ctk.CTkButton(
            content, text="Submit",
            fg_color=BRAND_ORANGE, hover_color=BRAND_PURPLE,
            command=close
        ).pack()

        # --- size & center behind main window ---
        popup.update_idletasks()
        w, h = popup.winfo_reqwidth(), popup.winfo_reqheight()
        screen_h = popup.winfo_screenheight()
        x_center = self.winfo_rootx() + (self.winfo_width() - w)//2
        y_center = self.winfo_rooty() + (self.winfo_height() - h)//2

        # first position off-screen, then reveal behind
        popup.geometry(f"{w}x{h}+{x_center}+{screen_h+50}")
        popup.deiconify()
        popup.lower(self)

        # then snap into place & grab focus
        def _snap_and_focus():
            popup.geometry(f"{w}x{h}+{x_center}+{y_center}")
            popup.attributes("-topmost", True)
            popup.grab_set()

        popup.after_idle(_snap_and_focus)
        popup.wait_window()

        return code_var.get().strip()


    def _copy_dropbox_code(self):
        """Copy the formatted Dropbox code to the clipboard."""
        code = self.output_entry.get()
        if code:
            # silently copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(code)
    
    def _generate_dropbox_code_threaded(self):
        """Threaded wrapper for dropbox code generation to prevent UI freezing."""
        dropbox_id = self.dropbox_id_entry.get().strip()
        driver_type = self.driver_type_var.get()
        reason_text = self.reason_var.get()
        custom_reason = ""

        if reason_text == "Select One:":
            messagebox.showerror("Missing Info", "Choose a reason.")
            return
        if reason_text == "Other":
            custom_reason = self.other_reason_entry.get().strip()
            if not custom_reason:
                messagebox.showerror("Missing Info", "Enter your custom reason.")
                return

        if driver_type == "Courier":
            employee_id = self.employee_id_entry.get().strip()
            if not employee_id:
                messagebox.showerror("Missing Info", "Enter Employee ID.")
                return
            driver_vendor_id = ""
            csa_number = ""
        else:
            employee_id = ""
            driver_vendor_id = self.driver_vendor_id_entry.get().strip()
            csa_number = self.csa_number_entry.get().strip()
            if not driver_vendor_id or not csa_number:
                messagebox.showerror("Missing Info", "Enter Vendor ID and CSA #.")
                return

        valid_monitors = [
            info for info in self.running_monitors.values()
            if info.get("monitor") and info["monitor"].driver
        ]
        if not valid_monitors:
            messagebox.showerror(
                "No Auto-Read",
                "No active Auto-Read session detected.\n"
                "Please start an Auto-Read monitor first."
            )
            return

        self.generate_button.configure(state="disabled", text="Generating...")
        self.output_entry.configure(state="normal")
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, "Working...")
        self.output_entry.configure(state="readonly")

        threading.Thread(
            target=self._generate_dropbox_code_worker,
            args=(dropbox_id, driver_type, employee_id, driver_vendor_id, csa_number, reason_text, custom_reason),
            daemon=True
        ).start()
    
    def _generate_dropbox_code_worker(self, dropbox_id, driver_type, employee_id, driver_vendor_id, csa_number, reason_text, custom_reason):
        """Background worker for dropbox code generation."""
        try:
            self._generate_dropbox_code_actual(dropbox_id, driver_type, employee_id, driver_vendor_id, csa_number, reason_text, custom_reason)
        except Exception as e:
            error_msg = str(e)
            self.after(0, lambda: self._update_dropbox_ui_error(error_msg))
    
    def _update_dropbox_ui_success(self, formatted_code):
        """Update UI after successful dropbox code generation (main thread)."""
        self.output_entry.configure(state="normal")
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, formatted_code)
        self.output_entry.configure(state="readonly")
        self.generate_button.configure(state="normal", text="Generate")
    
    def _update_dropbox_ui_error(self, error_message):
        """Update UI after dropbox code generation error (main thread)."""
        self.output_entry.configure(state="normal")
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, "Error occurred")
        self.output_entry.configure(state="readonly")
        self.generate_button.configure(state="normal", text="Generate")
        messagebox.showerror("Dropbox Error", f"Operation failed: {error_message}")
    
    def _generate_dropbox_code_actual(self, dropbox_id, driver_type, employee_id, driver_vendor_id, csa_number, reason_text, custom_reason):
        """Actual dropbox code generation logic (runs in background thread)."""
        from selenium.webdriver.support.ui import WebDriverWait, Select
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException
        from tkinter import messagebox

        # 1) Find an Auto-Read monitor with an active Selenium session
        valid_monitors = [
            info for info in self.running_monitors.values()
            if info.get("monitor") and info["monitor"].driver
        ]
        if not valid_monitors:
            self.after(0, lambda: messagebox.showerror(
                "No Auto-Read",
                "No active Auto-Read session detected.\n"
                "Please start an Auto-Read monitor first."
            ))
            self.after(0, lambda: self._update_dropbox_ui_error("No active Auto-Read session"))
            return

        # 2) Use the first valid monitor
        mon = valid_monitors[0]["monitor"]
        driver = mon.driver
        if driver is None:
            self.after(0, lambda: messagebox.showerror("Busy", "Auto-Read browser isn't ready yetâ€”please wait a moment."))
            self.after(0, lambda: self._update_dropbox_ui_error("Browser not ready"))
            return

        wait = WebDriverWait(driver, 20)

        # 3) Switch to existing Dropbox tab or open it if missing
        if self.dropbox_driver is driver and self.dropbox_handle in driver.window_handles:
            driver.switch_to.window(self.dropbox_handle)
            # Verify the tab is still on the correct page
            try:
                if "combination_search.php" not in driver.current_url:
                    print("ðŸ” Dropbox tab redirected - refreshing...", flush=True)
                    driver.get("https://internal.example.com")
                    wait.until(EC.presence_of_element_located((By.ID, "locid")))
            except Exception as e:
                print(f"âš ï¸ Dropbox tab validation failed: {e}", flush=True)
                # Recreate the tab
                driver.execute_script(
                    "window.open('https://internal.example.com','_blank');"
                )
                handle = driver.window_handles[-1]
                self.dropbox_driver = driver
                self.dropbox_handle = handle
                driver.switch_to.window(handle)
        else:
            driver.execute_script(
                "window.open('https://internal.example.com','_blank');"
            )
            handle = driver.window_handles[-1]
            self.dropbox_driver = driver
            self.dropbox_handle = handle
            driver.switch_to.window(handle)

        # â”€â”€â”€ If we got sent to dashboard.php, click the Combination Search link â”€â”€â”€
        if "dashboard.php" in driver.current_url:
            try:
                combo_link = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//a[@href='/combination_search.php']"
                )))
                combo_link.click()
                wait.until(EC.url_contains("combination_search.php"))
            except Exception:
                # fallback if link never appears (e.g. timeout / session expired)
                driver.get("https://internal.example.com")
                wait.until(EC.presence_of_element_located((By.ID, "locid")))

        try:
            # 4) Fill Dropbox ID and submit
            try:
                fld = wait.until(EC.element_to_be_clickable((By.ID, "locid")))
                fld.clear()
                fld.send_keys(dropbox_id, Keys.RETURN)
            except TimeoutException:
                # If element not found, try refreshing the page and retry once
                print("âš ï¸ Dropbox form not ready, refreshing page...", flush=True)
                driver.get("https://internal.example.com")
                wait.until(EC.presence_of_element_located((By.ID, "locid")))
                fld = wait.until(EC.element_to_be_clickable((By.ID, "locid")))
                fld.clear()
                fld.send_keys(dropbox_id, Keys.RETURN)

            # 5) Handle reason selection - either dropdown or textarea for "Other"
            try:
                if reason_text == "Other":
                    # For "Other", we need to select "Other" from dropdown first, then fill textarea
                    sel = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.ID, "reason"))
                    )
                    Select(sel).select_by_visible_text("Other")
                    
                    # Wait for the textarea to appear and fill it
                    other_textarea = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "reason_other"))
                    )
                    other_textarea.clear()
                    # Use the custom reason text passed from the main thread
                    other_textarea.send_keys(custom_reason)
                else:
                    # For predefined reasons, use dropdown selection
                    sel = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.ID, "reason"))
                    )
                    Select(sel).select_by_visible_text(reason_text)
            except TimeoutException:
                self.after(0, lambda: messagebox.showerror(
                    "Dropbox Error",
                    "Could not load the Reason dropdown in time.\n"
                    "Please try again in a moment."
                ))
                self.after(0, lambda: self._update_dropbox_ui_error("Reason dropdown timeout"))
                return

            # 6) Click driver type radio, then fill ID field(s)
            if driver_type == "Courier":
                courier_radio = wait.until(EC.element_to_be_clickable((By.ID, "courier_driver_courier")))
                courier_radio.click()
                emp = wait.until(EC.element_to_be_clickable((By.ID, "courier_number_ui")))
                emp.clear()
                emp.send_keys(employee_id, Keys.RETURN)
            else:
                driver_radio = wait.until(EC.element_to_be_clickable((By.ID, "courier_driver_driver")))
                driver_radio.click()
                vendor_el = wait.until(EC.element_to_be_clickable((By.ID, "driver_vendor_id")))
                vendor_el.clear()
                vendor_el.send_keys(driver_vendor_id)
                csa_el = wait.until(EC.element_to_be_clickable((By.ID, "csa_number")))
                csa_el.clear()
                csa_el.send_keys(csa_number, Keys.RETURN)

            # 7) Scrape the Manual Combination code
            p = wait.until(EC.presence_of_element_located((
                By.XPATH, "//p[contains(text(),'Manual Combination')]"
            )))
            raw_combo = p.text.split("Manual Combination:")[-1].strip()
            # break into digits to avoid double-commas, then re-join
            import re
            digits = re.findall(r"\d", raw_combo)
            formatted = "Dropbox Code - " + ", ".join(digits) + "."

            # 8) Update UI in main thread
            self.after(0, lambda: self._update_dropbox_ui_success(formatted))
            
            # 8a) Reset the Dropbox tab to the search page
            driver.get("https://internal.example.com")
            wait.until(EC.presence_of_element_located((By.ID, "locid")))

            # 8b) Clear the app's input fields for the next use
            self.after(0, lambda: self.dropbox_id_entry.delete(0, "end"))
            if driver_type == "Courier":
                self.after(0, lambda: self.employee_id_entry.delete(0, "end"))
            else:
                self.after(0, lambda: self.driver_vendor_id_entry.delete(0, "end"))
                self.after(0, lambda: self.csa_number_entry.delete(0, "end"))

        except TimeoutException as e:
            error_msg = str(e)
            self.after(0, lambda: messagebox.showerror(
                "Dropbox Timeout",
                f"Operation timed out while waiting for page elements.\n"
                f"This may be due to slow network or page loading issues.\n"
                f"Please try again in a moment.\n\n"
                f"Error: {error_msg}"
            ))
            self.after(0, lambda: self._update_dropbox_ui_error("Operation timed out"))
            return
        except Exception as e:
            error_msg = str(e)
            self.after(0, lambda: messagebox.showerror(
                "Dropbox Error",
                f"An unexpected error occurred: {error_msg}\n"
                f"Please try again or restart the Auto-Read monitor."
            ))
            self.after(0, lambda: self._update_dropbox_ui_error(error_msg))
            return

        finally:
            # 9) Leave the tab open (we track it), then switch back
            if self.dropbox_driver is not driver:
                # first time here, track this driver & handle
                self.dropbox_driver = driver
                self.dropbox_handle = driver.current_window_handle
            # always switch focus back to the main Auto-Read tab
            driver.switch_to.window(driver.window_handles[0])
    
    def _ensure_dropbox_tab(self):
        """
        Every 10 minutes: close the existing Dropbox tab and open a brand new one
        on the first Auto-Read driver to keep it fresh and active.
        Only runs when monitors are active.
        """
        # collect all live Auto-Read drivers
        drivers = [
            info["monitor"].driver
            for info in self.running_monitors.values()
            if info["monitor"].driver
        ]
        
        if not drivers:
            # no sessions â†’ clear tracking and don't schedule next check
            self.dropbox_driver = None
            self.dropbox_handle = None
            print("ðŸ” No active monitors - dropbox tab refresh stopped", flush=True)
            return
        
        # Only continue if we have active monitors
        driver = drivers[0]
        
        try:
            print("ðŸ”„ Closing existing dropbox tab and opening a new one...", flush=True)
            
            # Close the existing dropbox tab if it exists and is still valid
            if (self.dropbox_driver is driver and 
                self.dropbox_handle and 
                self.dropbox_handle in driver.window_handles):
                
                # Switch to the dropbox tab and close it
                driver.switch_to.window(self.dropbox_handle)
                driver.close()
                print("âœ… Closed existing dropbox tab", flush=True)
            
            # Switch back to the main tab (first tab)
            driver.switch_to.window(driver.window_handles[0])
            
            # Open a brand new dropbox tab
            driver.execute_script(
                "window.open('https://internal.example.com','_blank');"
            )
            self.dropbox_driver = driver
            self.dropbox_handle = driver.window_handles[-1]
            print("âœ… Opened new dropbox tab successfully", flush=True)
            
        except Exception as e:
            print(f"âš ï¸ Failed to refresh dropbox tab: {e}", flush=True)
            # Clear tracking on error
            self.dropbox_driver = None
            self.dropbox_handle = None
        
        # Only schedule next check if we have active monitors
        if drivers:
            # schedule next check in 10 minutes (600 seconds)
            self.after(600_000, self._ensure_dropbox_tab)
        else:
            # No active monitors - stop the refresh cycle
            self.dropbox_refresh_active = False

    def _start_dropbox_refresh_cycle(self):
        """
        Start the dropbox refresh cycle if not already active and monitors are running.
        This ensures the dropbox tab stays alive as long as any monitor is active.
        """
        # Check if we have active monitors
        drivers = [
            info["monitor"].driver
            for info in self.running_monitors.values()
            if info["monitor"].driver
        ]
        
        if not drivers:
            # No monitors running - don't start refresh cycle
            self.dropbox_refresh_active = False
            return
        
        # Only start if not already active
        if not self.dropbox_refresh_active:
            self.dropbox_refresh_active = True
            print("ðŸ”„ Starting dropbox refresh cycle (10-minute intervals)", flush=True)
            # Start the first refresh cycle
            self._ensure_dropbox_tab()

    def _on_driver_type_change(self, new_value: str):
        """Show Courier fields (Employee ID) or Service Provider Driver fields (Vendor ID, CSA #)."""
        if new_value == "Courier":
            self.employee_id_label.grid()
            self.employee_id_entry.grid()
            self.driver_vendor_label.grid_remove()
            self.driver_vendor_id_entry.grid_remove()
            self.csa_label.grid_remove()
            self.csa_number_entry.grid_remove()
        else:
            self.employee_id_label.grid_remove()
            self.employee_id_entry.grid_remove()
            self.driver_vendor_label.grid()
            self.driver_vendor_id_entry.grid()
            self.csa_label.grid()
            self.csa_number_entry.grid()

    def _on_reason_change(self, new_value: str):
        """Show the Other-Reason row when 'Other' is selected, and
        shift the Output row and Generate button up/down accordingly."""
        if new_value == "Other":
            self.other_reason_label.grid()
            self.other_reason_entry.grid()
            output_row = 5
            generate_row = 6
        else:
            self.other_reason_label.grid_remove()
            self.other_reason_entry.grid_remove()
            output_row = 4
            generate_row = 5

        self.output_label.grid_configure(row=output_row)
        self.output_entry.grid_configure(row=output_row)
        self.copy_button.grid_configure(row=output_row)
        self.generate_button.grid_configure(row=generate_row)

    def _breathing_animation(self):
        """Breathing animation for the running status"""
        current_text = self.sort_status_label.cget("text") if hasattr(self, 'sort_status_label') else ""
        
        # Check if we should animate (Running or Processing Stop Movement)
        if current_text in ["Running", "Processing Stop Movement"] or "Moving Disp#" in current_text:
            # Update alpha for breathing effect
            self.sort_breathing_alpha += 0.1 * self.sort_breathing_direction

            # Reverse direction at boundaries
            if self.sort_breathing_alpha <= 0.3:
                self.sort_breathing_direction = 1
            elif self.sort_breathing_alpha >= 1.0:
                self.sort_breathing_direction = -1

            # Apply breathing effect to text color
            import colorsys
            # Convert Carrier orange to RGB
            carrierx_orange_rgb = (255, 102, 0)  # #FF6600
            # Apply alpha
            new_color = tuple(int(c * self.sort_breathing_alpha) for c in carrierx_orange_rgb)
            # Convert back to hex
            hex_color = f"#{new_color[0]:02x}{new_color[1]:02x}{new_color[2]:02x}"

            self.sort_status_label.configure(text_color=hex_color)

            # Schedule next animation frame
            self.after(100, self._breathing_animation)

    # â”€â”€â”€ Auto-Sort start handler â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def start_sort(self):
        loc = self.sort_location_entry.get().strip()
        if not loc:
            from tkinter import messagebox
            messagebox.showerror("Missing Info", "Please enter a location to sort.")
            return

        # Check if user is logged in
        if not self.username or not self.password:
            from tkinter import messagebox
            messagebox.showerror("Not Logged In", "Please log in first before using auto-sort.")
            return

        # reset and show the progress bar now that we're starting
        self.update_custom_progress(0)

        # Clear results area and show "Running" status
        self.clear_results()
        self.sort_status_label.configure(text="Running")
        self.sort_breathing_alpha = 1.0
        self.sort_breathing_direction = -1
        self._breathing_animation()

        # Enable Stop button, disable Start button
        self.sort_stop_button.configure(state="normal")
        self.sort_start_button.configure(state="disabled")

        # Reset stop event
        self.sort_stop_event.clear()

        # launch the heavy work in a daemon thread
        threading.Thread(
            target=self._run_sort_worker,
            args=(loc,),
            daemon=True
        ).start()

    def draw_road_lines(self, progress_width=0):
        """Draw dashed road lines on the canvas - only show completed portion"""
        canvas_width = self.road_canvas.winfo_width()
        if canvas_width <= 1:  # Canvas not yet sized
            self.road_canvas.after(100, lambda: self.draw_road_lines(progress_width))
            return
            
        # Clear existing lines
        self.road_canvas.delete("road_line")
        
        # Only draw road lines up to the progress width
        y_center = 20
        dash_length = 20
        gap_length = 15
        
        x = 0
        while x < progress_width:
            # Calculate dash end position
            dash_end = min(x + dash_length, progress_width)
            
            # Only draw if dash is within progress area
            if x < progress_width:
                self.road_canvas.create_line(
                    x, y_center, dash_end, y_center,
                    fill=BRAND_ORANGE, width=2, tags="road_line"
                )
            x += dash_length + gap_length
    
    def update_custom_progress(self, percentage):
        """Update the custom progress bar with truck movement - thread-safe"""
        # Marshal to main thread if called from background thread
        if threading.current_thread() != threading.main_thread():
            self.after(0, lambda: self.update_custom_progress(percentage))
            return
        
        # Show truck when process starts (at 1% or higher)
        if percentage >= 1.0 and not self.truck_breathing_active:
            self.road_canvas.itemconfig(self.truck_id, state="normal")
            self.truck_breathing_active = True
            # Only start breathing animation when actual progress begins (above 1%)
            if percentage > 1.0:
                self.start_truck_breathing()
        
        # Hide truck when process stops
        if percentage == 0 and self.truck_breathing_active:
            self.road_canvas.itemconfig(self.truck_id, state="hidden")
            self.truck_breathing_active = False
            # Reset percentage text color to default when breathing stops
            self.progress_text.configure(text_color=BRAND_ORANGE)
        
        # Update progress text (only show when process is running)
        if percentage > 0:
            self.progress_text.configure(text=f"{percentage:.1f}%")
        else:
            self.progress_text.configure(text="")
        
        # Calculate progress width
        canvas_width = self.road_canvas.winfo_width()
        if canvas_width <= 1:  # Canvas not yet sized
            return
            
        # Calculate progress width (progressive - only show completed portion)
        progress_width = (percentage / 100) * (canvas_width - 10)  # Leave 5px padding on each side
        
        # Update progress canvas width (this creates the progressive effect)
        self.progress_canvas.configure(width=progress_width)
        
        # Move truck to the end of the progress (spearhead position)
        # Only move truck if progress is actually advancing
        if progress_width > 0:
            truck_x = progress_width + 5  # Position truck just ahead of the progress
            if truck_x < 10:  # Keep truck at minimum position
                truck_x = 10
            
            # Move truck - handle both image (2 coords) and rectangle (4 coords) trucks
            try:
                current_coords = self.road_canvas.coords(self.truck_id)
                if len(current_coords) == 2:  # Image truck
                    # For image truck, just set x, y coordinates
                    self.road_canvas.coords(self.truck_id, truck_x, 20)
                elif len(current_coords) == 4:  # Rectangle truck
                    # For rectangle truck, set x1, y1, x2, y2 coordinates
                    x1, y1, x2, y2 = current_coords
                    width = x2 - x1
                    height = y2 - y1
                    self.road_canvas.coords(self.truck_id, truck_x, y1, truck_x + width, y2)
            except Exception as e:
                # Silently handle any coordinate errors
                pass
        
        # Redraw road lines with current progress width
        self.draw_road_lines(progress_width)
    
    def start_truck_breathing(self):
        """Start the truck breathing animation"""
        if not self.truck_breathing_active:
            return
        
        # Update breathing alpha
        self.truck_breathing_alpha += self.truck_breathing_direction * 0.05
        
        # Reverse direction at boundaries
        if self.truck_breathing_alpha <= 0.3:
            self.truck_breathing_alpha = 0.3
            self.truck_breathing_direction = 1
        elif self.truck_breathing_alpha >= 1.0:
            self.truck_breathing_alpha = 1.0
            self.truck_breathing_direction = -1
        
        # Apply breathing effect to truck and percentage text
        try:
            # Get current truck position
            coords = self.road_canvas.coords(self.truck_id)
            if coords:
                # Handle both image (2 coords) and rectangle (4 coords) trucks
                if len(coords) == 2:  # Image truck
                    x, y = coords
                    # For image truck, create a breathing effect by slightly changing position
                    breathing_offset = (1.0 - self.truck_breathing_alpha) * 2  # Subtle movement
                    self.road_canvas.coords(self.truck_id, x + breathing_offset, y)
                elif len(coords) == 4:  # Rectangle truck
                    x1, y1, x2, y2 = coords
                    # For rectangle truck, change the fill color intensity
                    color_intensity = int(255 * self.truck_breathing_alpha)
                    breathing_color = f"#{color_intensity:02x}6600"  # Orange with varying intensity
                    self.road_canvas.itemconfig(self.truck_id, fill=breathing_color)
            
            # Apply breathing effect to percentage text
            # Calculate color intensity for the percentage text
            text_color_intensity = int(255 * self.truck_breathing_alpha)
            breathing_text_color = f"#{text_color_intensity:02x}6600"  # Orange with varying intensity
            
            # Update percentage text color to breathe in sync
            self.progress_text.configure(text_color=breathing_text_color)
            
        except Exception as e:
            # Silently handle any coordinate errors
            pass
        
        # Schedule next frame
        if self.truck_breathing_active:
            self.after(50, self.start_truck_breathing)  # 20 FPS animation

    def clear_results(self, show_select_all=True):
        """Clear all result frames and hide bottom box"""
        # If UI doesn't exist yet, just clear stored results
        if not hasattr(self, 'result_frames'):
            if hasattr(self, 'stored_results'):
                self.stored_results.clear()
            return
        

            
        for result in self.result_frames:
            result['frame'].destroy()
        self.result_frames.clear()
        # Reset select all checkbox (only if it exists)
        if hasattr(self, 'select_all_var'):
            self.select_all_var.set(False)
        # Show select all checkbox again (in case it was hidden) - unless explicitly told not to
        if show_select_all and hasattr(self, 'select_all_checkbox'):
            self.select_all_checkbox.pack(anchor="w", padx=5, pady=(5,10))
        # Hide the bottom box (only if it exists)
        if hasattr(self, 'bottom_box'):
            self.bottom_box.pack_forget()
        
        # Clear stored results as well
        if hasattr(self, 'stored_results'):
            self.stored_results.clear()
        
        # Scroll to top after clearing results
        self.scroll_results_to_top()
    
    def add_result(self, text, show_checkbox=True):
        """Add a new result frame with checkbox and label"""
        # If results_scrollable doesn't exist yet, store the result for later
        if not hasattr(self, 'results_scrollable'):
            if not hasattr(self, 'stored_results'):
                self.stored_results = []
            self.stored_results.append(text)
            print(f"ðŸ” Stored result for later: {text[:50]}...")
            return
        
        print(f"ðŸ” Adding result to UI: {text[:50]}...")
        
        # Create a frame to hold checkbox and label
        frame = ctk.CTkFrame(self.results_scrollable, fg_color="transparent")
        frame.pack(anchor="w", padx=5, pady=1, fill="x")
        
        # Create checkbox variable
        checkbox_var = tk.BooleanVar()
        
        # Create checkbox (only for mismatch results and if show_checkbox is True)
        if text.strip().startswith("âŒ") and show_checkbox:
            print(f"ðŸ” Creating checkbox for: {text[:50]}... (show_checkbox={show_checkbox})")
            checkbox = ctk.CTkCheckBox(
                frame,
                text="",
                variable=checkbox_var,
                command=self.update_select_all_state,
                fg_color=BRAND_ORANGE,
                hover_color=BRAND_PURPLE,
                width=20
            )
            checkbox.pack(side="left", padx=(0, 5))
            

        else:
            print(f"ðŸ” No checkbox for: {text[:50]}... (show_checkbox={show_checkbox}, starts_with_âŒ={text.strip().startswith('âŒ')})")
            # For non-mismatch results or when checkboxes are hidden, create an empty space
            spacer = ctk.CTkLabel(frame, text="", width=20)
            spacer.pack(side="left", padx=(0, 5))
            checkbox = None
        
        # Create label
        label = ctk.CTkLabel(
            frame,
            text=text,
            text_color=BRAND_ORANGE,
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left"
        )
        label.pack(side="left", fill="x", expand=True)
        
        # Store the frame, checkbox, and label
        self.result_frames.append({
            'frame': frame,
            'checkbox': checkbox,
            'checkbox_var': checkbox_var,
            'label': label,
            'text': text
        })
        
        # Scroll to top after adding new result
        self.scroll_results_to_top()

    def scroll_results_to_top(self):
        """Scroll the results scrollable frame to the top"""
        try:
            if hasattr(self, 'results_scrollable') and self.results_scrollable:
                # Access the internal canvas and scroll to top
                self.results_scrollable._parent_canvas.yview_moveto(0)
        except Exception as e:
            print(f"ðŸ” Error scrolling to top: {e}")

    def hide_checkboxes(self):
        """Hide all checkboxes in the results and select all checkbox"""
        print("ðŸ” Hiding checkboxes...")
        # Hide select all checkbox (only if it exists)
        if hasattr(self, 'select_all_checkbox'):
            self.select_all_checkbox.pack_forget()
            print("ðŸ” Select all checkbox hidden")
        else:
            print("ðŸ” Select all checkbox doesn't exist yet")
        
        # Hide individual checkboxes
        for result in self.result_frames:
            if result['checkbox'] is not None:
                result['checkbox'].pack_forget()
                # Create a spacer to maintain alignment
                spacer = ctk.CTkLabel(result['frame'], text="", width=20)
                spacer.pack(side="left", padx=(0, 5))
                result['checkbox'] = None
        print("ðŸ” All checkboxes hidden")

    def toggle_select_all(self):
        """Toggle all mismatch checkboxes based on select all checkbox"""
        if not hasattr(self, 'select_all_var'):
            return
        select_all = self.select_all_var.get()
        for result in self.result_frames:
            if result['checkbox'] is not None:  # Only mismatch results have checkboxes
                result['checkbox_var'].set(select_all)

    def update_select_all_state(self):
        """Update select all checkbox state based on individual checkboxes"""
        mismatch_results = [r for r in self.result_frames if r['checkbox'] is not None]
        if not mismatch_results:
            return
        
        checked_count = sum(1 for r in mismatch_results if r['checkbox_var'].get())
        total_count = len(mismatch_results)
        
        if hasattr(self, 'select_all_var'):
            if checked_count == 0:
                self.select_all_var.set(False)
            elif checked_count == total_count:
                self.select_all_var.set(True)
        else:
            # Partially selected - don't change select all state
            pass

    def copy_results_to_clipboard(self):
        """Copy selected results to clipboard"""
        try:
            # Check if there are any results to copy
            if not hasattr(self, 'result_frames') or not self.result_frames:
                messagebox.showinfo("Nothing to Copy", "No results available to copy.")
                return
            
            # Gather selected result text
            all_text = []
            # Get the location from the sort location entry
            location = self.sort_location_entry.get().strip().upper()
            if location:
                all_text.append(f"=== Route Mismatch Results for {location} ===")
            else:
                all_text.append("=== Route Mismatch Results ===")
            all_text.append("")
            
            selected_count = 0
            total_mismatches = 0
            
            for result in self.result_frames:
                text = result['text']
                if text:  # Only add non-empty text
                    # Check if this is a mismatch result and if it's selected
                    if text.strip().startswith("âŒ"):
                        total_mismatches += 1
                        if result['checkbox'] is not None and result['checkbox_var'].get():
                            all_text.append(text)
                            selected_count += 1
                    else:
                        # Always include non-mismatch text (headers, summaries, etc.)
                        all_text.append(text)
            
            # If no mismatches are selected, ask user if they want to copy all
            if total_mismatches > 0 and selected_count == 0:
                response = messagebox.askyesno("No Selection", 
                    f"No mismatch results are selected. Would you like to copy all {total_mismatches} results?")
                if response:
                    # Copy all mismatches
                    all_text = []
                    # Get the location from the sort location entry
                    location = self.sort_location_entry.get().strip().upper()
                    if location:
                        all_text.append(f"=== Route Mismatch Results for {location} ===")
                    else:
                        all_text.append("=== Route Mismatch Results ===")
                    all_text.append("")
                    
                    for result in self.result_frames:
                        text = result['text']
                        if text:
                            all_text.append(text)
                    selected_count = total_mismatches
                else:
                    return
            
            # Join all text with newlines
            clipboard_text = "\n".join(all_text)
            
            # Copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(clipboard_text)
            
            # Show success message with selected count
            if selected_count > 0:
                messagebox.showinfo("Copied!", f"Copied {selected_count} mismatch results to clipboard.")
            else:
                messagebox.showinfo("Copied!", "Copied summary information to clipboard.")
            
        except Exception as e:
            messagebox.showerror("Copy Error", f"Failed to copy results: {e}")

    def stop_sort(self):
        """Stop the auto-sort process by killing the Selenium process"""
        # Set the stop event
        self.sort_stop_event.set()
        
        # Update UI
        self.sort_status_label.configure(text="Stopping...", text_color="red")
        self.sort_stop_button.configure(state="disabled")
        
        # Kill the Selenium process immediately
        try:
            if hasattr(self, 'driver') and self.driver:
                self.driver.quit()
                self.driver = None
                print("Selenium driver terminated")
        except Exception as e:
            print(f"Error terminating driver: {e}")
        
        # Kill any remaining Chrome processes
        try:
            import psutil
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'chrome' in proc.info['name'].lower() and 'MessageClear' in ' '.join(proc.cmdline()).lower():
                        proc.terminate()
                        print(f"Terminated Chrome process: {proc.info['pid']}")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
        except Exception as e:
            print(f"Error cleaning up processes: {e}")
        
        # Update status and reset progress
        self.sort_status_label.configure(text="Stopped")
        self.update_custom_progress(0)  # This will hide the percentage text
        
        # Re-enable start button immediately
        self.sort_start_button.configure(state="normal")
        self.sort_stop_button.configure(state="disabled")

    def _sso_navigate_sort(self, driver, wait,
                      url="https://internal.example.com",
                      username=None, password=None):
        """
        Navigate to `url` and, if you land on the SSO/Identity Portal page,
        submit creds and wait for the FRO menu to appear.
        """
        logging.info("ðŸ”„ Navigating to DSM...")
        driver.get(url)

        # â€” Poll frequently for auto-login (automatic login via SSO/cookies can take ~10 seconds) â€”
        logging.info("ðŸ” Checking for auto-login (polling every 1 second for up to 15 seconds)...")
        for attempt in range(15):  # Check for up to 15 seconds
            if CarrierSeleniumMonitor.check_if_logged_in(driver):
                logging.info(f"âœ… Auto-login detected after {attempt + 1} seconds (URL contains /home and FRO menu found)")
                return  # Already logged in, skip login process
            time.sleep(1)
        
        # Final check after polling period
        if CarrierSeleniumMonitor.check_if_logged_in(driver):
            logging.info("âœ… Auto-login detected after polling period")
            return  # Already logged in, skip login process

        # Check if we need to log in or if already logged in
        try:
            logging.info("ðŸ”„ Checking for login form...")
            # Use longer timeout for login detection (headless needs more time)
            login_wait = WebDriverWait(driver, 30)
            
            # Wait for username input field to appear using dynamic detection (with frequent auto-login checks)
            logging.info("ðŸ” Waiting for username input field to appear...")
            user = None
            start_time = time.time()
            timeout = 30
            check_interval = 1  # Check every 1 second
            
            while time.time() - start_time < timeout:
                # Check for auto-login first
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected during username field wait")
                    return  # Already logged in, skip login process
                
                # Try to find username field
                try:
                    user = CarrierSeleniumMonitor.find_username_input(driver, None)
                    if user and user.is_displayed():
                        break
                except:
                    pass
                
                time.sleep(check_interval)
            
            if not user:
                # Before raising error, final check for auto-login
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected (final check)")
                    return  # Already logged in, skip login process
                
                # Fallback: try original method
                try:
                    user = login_wait.until(CarrierSeleniumMonitor.UsernameFieldFound())
                except TimeoutException:
                    # Final check for auto-login before raising exception
                    if CarrierSeleniumMonitor.check_if_logged_in(driver):
                        logging.info("âœ… Auto-login detected during username wait timeout")
                        return  # Already logged in, skip login process
                    
                    # Fallback to waiting for any clickable input (original behavior)
                    logging.warning("âš ï¸ Could not dynamically find username input, trying fallback...")
                    try:
                        # Try common input ID patterns
                        for input_id in ["input44", "input28", "input"]:
                            try:
                                user = login_wait.until(EC.element_to_be_clickable((By.ID, input_id)))
                                break
                            except:
                                continue
                        # If still not found, get first visible text input
                        if not user:
                            user = login_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
                    except:
                        # Final check for auto-login before raising exception
                        if CarrierSeleniumMonitor.check_if_logged_in(driver):
                            logging.info("âœ… Auto-login detected (fallback check)")
                            return  # Already logged in, skip login process
                        raise Exception("Could not find username input field")
            
            # Login form found - perform two-step login
            logging.info("ðŸ”„ Login form detected - performing two-step login...")
            logging.info(f"âœ… Found username input: {user.get_attribute('id') or user.get_attribute('name') or 'unknown'}")
            
            # Step 1: Enter username and click Next
            logging.info("ðŸ”„ Step 1: Entering username...")
            user.clear()
            user.send_keys(username)
            
            # Look for Next button - try multiple possible selectors
            next_button = None
            next_selectors = [
                "input[type='submit']",
                "button[type='submit']", 
                "input[value*='Next']",
                "button:contains('Next')",
                ".next-button",
                "#next-button",
                "input[value='Next']",
                "button[value='Next']"
            ]
            
            for selector in next_selectors:
                try:
                    if selector.startswith("button:contains"):
                        # Use XPath for text content
                        next_button = login_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]")))
                    else:
                        next_button = login_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    break
                except TimeoutException:
                    continue
            
            if next_button:
                logging.info("ðŸ”„ Found Next button, clicking...")
                next_button.click()
                # Wait a moment for the page to transition
                time.sleep(2)
            else:
                logging.info("ðŸ”„ No Next button found, trying Enter key...")
                user.send_keys(Keys.RETURN)
                time.sleep(2)
            
            # Step 2: Enter password
            logging.info("ðŸ”„ Step 2: Entering password...")
            # Wait for password input field to appear using dynamic detection (with frequent auto-login checks)
            logging.info("ðŸ” Waiting for password input field to appear...")
            pw = None
            start_time = time.time()
            timeout = 30
            check_interval = 1  # Check every 1 second
            
            while time.time() - start_time < timeout:
                # Check for auto-login first
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected during password field wait")
                    return  # Already logged in, skip login process
                
                # Try to find password field
                try:
                    pw = CarrierSeleniumMonitor.find_password_input(driver, None)
                    if pw and pw.is_displayed():
                        break
                except:
                    pass
                
                time.sleep(check_interval)
            
            if not pw:
                # Before raising error, final check for auto-login
                if CarrierSeleniumMonitor.check_if_logged_in(driver):
                    logging.info("âœ… Auto-login detected (final check)")
                    return  # Already logged in, skip login process
                
                # Fallback: try original method
                try:
                    pw = login_wait.until(CarrierSeleniumMonitor.PasswordFieldFound())
                except TimeoutException:
                    # Final check for auto-login before raising exception
                    if CarrierSeleniumMonitor.check_if_logged_in(driver):
                        logging.info("âœ… Auto-login detected during password wait timeout")
                        return  # Already logged in, skip login process
                    
                    # Fallback to waiting for password input (original behavior)
                    logging.warning("âš ï¸ Could not dynamically find password input, trying fallback...")
                    try:
                        # Try common input ID patterns
                        for input_id in ["input70", "input54", "input"]:
                            try:
                                pw = login_wait.until(EC.element_to_be_clickable((By.ID, input_id)))
                                break
                            except:
                                continue
                        # If still not found, get password type input
                        if not pw:
                            pw = login_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
                    except:
                        # Final check for auto-login before raising exception
                        if CarrierSeleniumMonitor.check_if_logged_in(driver):
                            logging.info("âœ… Auto-login detected (fallback check)")
                            return  # Already logged in, skip login process
                        raise Exception("Could not find password input field")
            
            logging.info(f"âœ… Found password input: {pw.get_attribute('id') or pw.get_attribute('name') or 'unknown'}")
            pw.clear()
            pw.send_keys(password + Keys.RETURN)
            
            # Wait for DSM to be fully loaded after login
            logging.info("ðŸ”„ Waiting for DSM to load after login...")
            self._wait_for_dsm_ready(driver, wait)
            
        except TimeoutException:
            # No login form found - check if already logged in
            logging.info("ðŸ”„ No login form found - checking if already logged in...")
            self._wait_for_dsm_ready(driver, wait)
    
    def _wait_for_dsm_ready(self, driver, wait):
        """
        Wait for DSM to be fully loaded and ready to use.
        """
        logging.info("ðŸ”„ Checking for DSM page elements...")
        
        # Check if DSM is ready by looking for key elements
        def dsm_page_ready(driver):
            try:
                # Check for FRO menu (main navigation)
                fro_menu = driver.find_element(By.XPATH, "//span[normalize-space(text())='FRO']")
                
                # Check for other key DSM elements
                # Look for any DSM-specific elements that indicate the page is loaded
                dsm_elements = driver.find_elements(By.XPATH, "//span[contains(text(), 'FRO') or contains(text(), 'Pickup') or contains(text(), 'Message')]")
                
                # Check if elements are interactive
                return (fro_menu.is_enabled() and len(dsm_elements) > 0)
            except:
                return False
        
        try:
            # Wait for DSM to be fully ready (headless needs more time)
            # Use a longer timeout specifically for DSM page loading
            dsm_wait = WebDriverWait(driver, 90)  # 90 seconds for DSM page loading
            
            # First, wait for the page to be fully loaded
            logging.info("ðŸ”„ Waiting for DSM page to fully load...")
            time.sleep(5)  # Give extra time for JavaScript to execute
            
            # Then wait for elements to become interactive
            dsm_wait.until(dsm_page_ready)
            logging.info("âœ… DSM page loaded successfully - all elements interactive")
        except TimeoutException:
            logging.error("âŒ DSM page failed to load - elements not interactive")
            raise

    def _run_sort_worker(self, loc):
        # â”€â”€â”€ Import required libraries â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        import pandas as pd
        import logging
        import threading
        
        # â”€â”€â”€ Initialize headless Chrome for Auto-Sort â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        driver_path = ChromeDriverManager().install()
        service     = Service(driver_path, log_path=os.devnull)
        opts        = Options()
        opts.add_argument("--disable-gpu")
        opts.add_argument("--headless")  # Run in headless mode
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--log-level=3")
        # Performance optimizations
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-plugins")
        opts.add_argument("--disable-images")
        opts.add_argument("--disable-css")
        opts.add_argument("--disable-animations")
        opts.add_argument("--disable-web-security")
        opts.add_argument("--disable-features=VizDisplayCompositor")
        opts.add_argument("--memory-pressure-off")
        opts.add_argument("--max_old_space_size=4096")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        # â†’ ensure downloads go to ./downloads without prompt
        download_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(download_dir, exist_ok=True)
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        opts.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(service=service, options=opts)
        self.driver = driver  # Store reference for stop function
        wait   = WebDriverWait(driver, 30)

        # 1) Go to DSM homepage & perform SSO login with MFA callbacks
        dsm_url = "https://internal.example.com"
        
        # Use the same login approach as other functions - with MFA callbacks
        login_helper = CarrierSeleniumMonitor(
            username=self.username,
            password=self.password,
            dispatch_location=self.sort_location_entry.get().strip().upper(),
            stop_event=threading.Event(),
            message_types=[],
            mfa_factor_callback=self.prompt_mfa_factor,
            mfa_code_callback=self.prompt_mfa_code
        )
        
        # Navigate and login using the proven method
        login_helper.sso_navigate(
            driver, wait,
            dsm_url,
            self.username, self.password
        )
        
        print("âœ… Login successful using existing MFA callback system!")
        # â”€â”€â”€ STORE THE DSM TAB HANDLE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.dsm_handle = driver.current_window_handle

        # 2) Open SIMS in a new tab
        sims_url = "https://internal.example.com"
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        # â”€â”€â”€ STORE THE SIMS TAB HANDLE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.sims_handle = driver.current_window_handle
        driver.get(sims_url)

        # 3) On SIMS: enter location, select "All" dispatch type, then start download
        logging.info("ðŸ”„ Waiting for SIMS page to load...")
        
        # Wait for page to be ready by checking multiple elements
        sims_wait = WebDriverWait(driver, 15)  # Reasonable timeout
        try:
            # Wait for the page to be fully loaded and interactive
            logging.info("ðŸ”„ Checking for SIMS page elements...")
            
            # Check if page is ready by looking for the main form elements
            def sims_page_ready(driver):
                try:
                    # Check for all required elements
                    loc_input = driver.find_element(By.NAME, "searchAddressSiRequest.locationId")
                    dispatch_sel = driver.find_element(By.NAME, "searchAddressSiRequest.dispatchType")
                    excel_btn = driver.find_element(By.XPATH, "//input[@name='action' and @value='Excel Report']")
                    
                    # Check if elements are actually interactive (not just present)
                    return (loc_input.is_enabled() and 
                           dispatch_sel.is_enabled() and 
                           excel_btn.is_enabled())
                except:
                    return False
            
            # Wait for the page to be fully ready
            sims_wait.until(sims_page_ready)
            logging.info("âœ… SIMS page loaded successfully - all elements interactive")
        except TimeoutException:
            logging.error("âŒ SIMS page failed to load - elements not interactive")
            raise
        
        # Now get the location input field
        loc_input = driver.find_element(By.NAME, "searchAddressSiRequest.locationId")
        loc_input.clear()
        loc_input.send_keys(loc)
        logging.info(f"âœ… Entered location: {loc}")

        # Set dispatch type to "All"
        try:
            dispatch_sel = Select(driver.find_element(By.NAME, "searchAddressSiRequest.dispatchType"))
            dispatch_sel.select_by_value("AL")
            logging.info("âœ… Set dispatch type to 'All'")
        except Exception as e:
            logging.error(f"âŒ Failed to set dispatch type: {e}")
            raise

        # Find Excel download button
        try:
            excel_btn = driver.find_element(
                By.XPATH, "//input[@name='action' and @value='Excel Report']"
            )
            logging.info("âœ… Found Excel download button")
        except Exception as e:
            logging.error(f"âŒ Failed to find Excel download button: {e}")
            raise
        
        # record existing files so we only pick up the new download
        initial_files = set(os.listdir(download_dir))
        
        # Start the Excel download
        logging.info("ðŸ”„ Starting SIMS Excel download...")
        from urllib3.exceptions import ReadTimeoutError
        
        # Start progress bar at 1% when SIMS download begins
        self.update_custom_progress(1.0)
        
        download_started = False
        for attempt in range(1, 4):
            try:
                # Re-find the Excel button on each attempt to avoid stale element
                excel_btn = driver.find_element(
                    By.XPATH, "//input[@name='action' and @value='Excel Report']"
                )
                excel_btn.click()
                logging.info("ðŸ“¥ SIMS Excel download started - switching to DSM while it downloads...")
                download_started = True
                break
            except ReadTimeoutError:
                logging.warning(f"ReadTimeoutError on excel_btn.click() â€“ retry {attempt}/3")
                time.sleep(2)
            except Exception as e:
                logging.warning(f"Error on excel_btn.click() â€“ retry {attempt}/3: {e}")
                time.sleep(2)
        else:
            # if all retries failed, raise a proper error
            raise RuntimeError("Failed to start SIMS Excel download after 3 attempts")

        # 4) Switch back to DSM tab while SIMS downloads in background
        driver.switch_to.window(self.dsm_handle)
        
        # Navigate to Pickup List while SIMS downloads
        logging.info("ðŸ”„ Navigating to DSM Pickup List while SIMS downloads...")
        
        # Wait for DSM to be fully loaded
        time.sleep(2)
        
        # Navigate to Pickup List with better error handling
        try:
            pickup_list_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(text(), 'Pickup List')]")
            ))
            pickup_list_btn.click()
            logging.info("âœ… Clicked Pickup List button")
        except TimeoutException:
            # Try alternative selectors
            try:
                pickup_list_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//span[normalize-space(text())='Pickup List']")
                ))
                pickup_list_btn.click()
                logging.info("âœ… Clicked Pickup List button (alternative selector)")
            except TimeoutException:
                # Try clicking via FRO menu first
                try:
                    fro_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='FRO']")))
                    fro_btn.click()
                    time.sleep(1)
                    pickup_list_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[normalize-space(text())='Pickup List']")
                    ))
                    pickup_list_btn.click()
                    logging.info("âœ… Clicked Pickup List via FRO menu")
                except Exception as e:
                    logging.error(f"âŒ Could not navigate to Pickup List: {e}")
                    raise
        
        # Wait for the page to load
        time.sleep(3)
        
        # Set dispatch location
        try:
            fld = wait.until(EC.element_to_be_clickable((By.ID, "dispatchLocation")))
            fld.clear()
            fld.send_keys(loc + Keys.RETURN)
            logging.info(f"âœ… Set dispatch location to {loc}")
        except TimeoutException:
            # Try alternative selector
            try:
                fld = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[formcontrolname='dispatchLocation']")
                ))
                fld.clear()
                fld.send_keys(loc + Keys.RETURN)
                logging.info(f"âœ… Set dispatch location to {loc} (alternative selector)")
            except Exception as e:
                logging.error(f"âŒ Could not set dispatch location: {e}")
                raise
        
        # Wait for table to load
        time.sleep(3)
        
        logging.info(f"âœ… DSM Pickup List loaded for location {loc}")
        
        # Sort out regular pickups by entering "REG" into the P/U Type field
        logging.info("ðŸ”„ Sorting regular pickups by entering 'REG' into P/U Type field...")
        try:
            # Find the P/U Type input field using multiple specific selectors
            pu_type_input = None
            
            # Method 1: Target the specific class containing "P/U Type"
            try:
                pu_type_input = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "input[class*='P/U Type']"
                )))
                logging.info("âœ… Found P/U Type input field (method 1: class contains 'P/U Type')")
            except:
                pass
            
            # Method 2: Use XPath to find input with class containing "P/U Type"
            if pu_type_input is None:
                try:
                    pu_type_input = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[contains(@class, 'P/U Type')]"
                    )))
                    logging.info("âœ… Found P/U Type input field (method 2: XPath with class)")
                except:
                    pass
            
            # Method 3: Look for input with name='tab' and class containing "P/U Type"
            if pu_type_input is None:
                try:
                    pu_type_input = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@name='tab' and contains(@class, 'P/U Type')]"
                    )))
                    logging.info("âœ… Found P/U Type input field (method 3: name='tab' + class)")
                except:
                    pass
            
            # Method 4: Fallback to original selector
            if pu_type_input is None:
                try:
                    pu_type_input = wait.until(EC.element_to_be_clickable((
                        By.CSS_SELECTOR, "input[name='tab'][type='text'][placeholder='']"
                    )))
                    logging.info("âœ… Found P/U Type input field (method 4: fallback selector)")
                except:
                    pass
            
            if pu_type_input is not None:
                pu_type_input.clear()
                pu_type_input.send_keys("REG")
                logging.info("âœ… Entered 'REG' into P/U Type field to sort regular pickups")
                
                # Wait a moment for the filter to apply
                time.sleep(2)
            else:
                logging.warning("âš ï¸ Could not find P/U Type input field with any method")
                
        except Exception as e:
            logging.warning(f"âš ï¸ Could not enter 'REG' into P/U Type field: {e}")
            logging.warning("Continuing with CSV download without sorting...")
        
        # Download DSM Pickup List as CSV
        logging.info("ðŸ”„ Downloading DSM Pickup List as CSV...")
        try:
            # Find CSV download button with multiple fallback methods
            csv_btn = None
            csv_selectors = [
                "//span[@class='p-button-label' and text()='Download - CSV']",
                "//span[contains(text(), 'Download - CSV')]",
                "//button[contains(text(), 'Download - CSV')]",
                "//a[contains(text(), 'Download - CSV')]",
                "//span[contains(text(), 'CSV')]",
                "//button[contains(text(), 'CSV')]",
                "//a[contains(text(), 'CSV')]",
                "//*[contains(text(), 'Download - CSV')]",
                "//*[contains(text(), 'CSV')]"
            ]
            
            for i, selector in enumerate(csv_selectors, 1):
                try:
                    csv_btn = driver.find_element(By.XPATH, selector)
                    logging.info(f"âœ… Found DSM CSV download button (method {i}): {selector}")
                    break
                except Exception as e:
                    logging.warning(f"âš ï¸ DSM CSV method {i} failed: {selector} - {e}")
                    continue
            
            if csv_btn is None:
                logging.warning("âš ï¸ Could not find DSM CSV download button, continuing without CSV download")
            else:
                # Record existing files before download
                dsm_initial_files = set(os.listdir(download_dir))
                
                # Click CSV download button
                csv_btn.click()
                logging.info("ðŸ“¥ DSM CSV download started")
                
                # Wait for CSV download to complete
                csv_timeout = 30
                csv_end_time = time.time() + csv_timeout
                dsm_csv_file = None
                
                while time.time() < csv_end_time:
                    csv_candidates = [
                        f for f in os.listdir(download_dir)
                        if f not in dsm_initial_files and f.lower().endswith(".csv")
                    ]
                    if csv_candidates:
                        dsm_csv_file = max(
                            csv_candidates,
                            key=lambda f: os.path.getctime(os.path.join(download_dir, f))
                        )
                        if not any(dsm_csv_file.endswith(ext) for ext in (".crdownload", ".part")):
                            break
                    time.sleep(1)
                
                if dsm_csv_file:
                    dsm_csv_path = os.path.join(download_dir, dsm_csv_file)
                    self.dsm_df = pd.read_csv(dsm_csv_path)
                    logging.info(f"ðŸ“Š Loaded DSM CSV report: {dsm_csv_path} (rows={len(self.dsm_df)})")
                    logging.info(f"DSM CSV columns: {list(self.dsm_df.columns)}")
                    
                    # Debug: Print first few rows to see the data structure
                    print(f"DSM CSV columns: {list(self.dsm_df.columns)}")
                    print("First 3 rows of DSM data:")
                    print(self.dsm_df.head(3))
                else:
                    logging.warning("âš ï¸ DSM CSV download timed out or failed")
                    
        except Exception as e:
            logging.warning(f"âš ï¸ Error downloading DSM CSV: {e}")
            logging.warning("Continuing with manual pickup list processing...")

        # 5) Now wait for SIMS download to complete
        logging.info("â³ Waiting for SIMS download to complete...")
        timeout = 120
        end_time = time.time() + timeout
        latest_file = None
        while time.time() < end_time:
            # only consider newly added .xls/.xlsx files
            candidates = [
                f for f in os.listdir(download_dir)
                if f not in initial_files and f.lower().endswith((".xls", ".xlsx"))
            ]
            if candidates:
                # pick the most recent among the new candidates
                latest_file = max(
                    candidates,
                    key=lambda f: os.path.getctime(os.path.join(download_dir, f))
                )
                # ensure it's not still downloading (handles both Chrome & other temp extensions)
                if not any(latest_file.endswith(ext) for ext in (".crdownload", ".part")):
                    break
            time.sleep(1)
        else:
            raise RuntimeError("ðŸ“¥ Download timed out after 120s")

        file_path = os.path.join(download_dir, latest_file)

        # â”€â”€â”€ Load into pandas for comparison later â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self.sims_df = pd.read_excel(file_path, header=2)
        logging.info(f"ðŸ“Š Loaded SIMS report: {file_path} (rows={len(self.sims_df)})")

        # â”€â”€â”€ Static mapping for the SIMS columns we actually compare â”€â”€â”€â”€â”€â”€â”€â”€â”€
        column_mapping = {
            "Account":      "Acc#/Loc Id",
            "Customer":     "Co Name",
            "Ready":        "Rdy Tm",
            "Close":        "Cls Tm",
            "Asg Route":    "Primary Route",
            "Street":       "Street",            # â† add Street mapping
            "Monthly/Day of Month": "Monthly/Day of Month",  # â† add Monthly/Day of Month mapping
        }
        
        # Debug: Print available columns
        print(f"Available SIMS columns: {list(self.sims_df.columns)}")
        print(f"First few rows of SIMS data:")
        print(self.sims_df.head())
        # verify they all exist
        missing = [hdr for hdr in column_mapping.values() if hdr not in self.sims_df.columns]
        if missing:
            raise RuntimeError(
                f"Missing expected columns in SIMS report: {missing}\n"
                f"Available columns: {list(self.sims_df.columns)}"
            )
        # DSM navigation already completed above while SIMS was downloading

        # â”€â”€â”€ Detect header column positions with multiple fallback methods â”€â”€â”€â”€â”€â”€â”€â”€
        header_names = []
        header_cells = []
        
        # Method 1: Try the original selector
        try:
            header_cells = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "tr.ng-star-inserted > th")
            ))
            header_names = [cell.text.strip() for cell in header_cells if cell.text.strip()]
            logging.info(f"âœ… Found headers (method 1): {header_names}")
        except TimeoutException:
            logging.warning("âš ï¸ Method 1 failed, trying alternative selectors")
            
            # Method 2: Try different CSS selector
            try:
                header_cells = wait.until(EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "thead th")
                ))
                header_names = [cell.text.strip() for cell in header_cells if cell.text.strip()]
                logging.info(f"âœ… Found headers (method 2): {header_names}")
            except TimeoutException:
                logging.warning("âš ï¸ Method 2 failed, trying XPath")
                
                # Method 3: Try XPath
                try:
                    header_cells = wait.until(EC.presence_of_all_elements_located(
                        (By.XPATH, "//thead//th")
                    ))
                    header_names = [cell.text.strip() for cell in header_cells if cell.text.strip()]
                    logging.info(f"âœ… Found headers (method 3): {header_names}")
                except TimeoutException:
                    logging.warning("âš ï¸ Method 3 failed, trying table headers")
                    
                    # Method 4: Try any table header
                    try:
                        header_cells = driver.find_elements(By.TAG_NAME, "th")
                        header_names = [cell.text.strip() for cell in header_cells if cell.text.strip()]
                        logging.info(f"âœ… Found headers (method 4): {header_names}")
                    except Exception as e:
                        logging.error(f"âŒ All header detection methods failed: {e}")
                        raise RuntimeError("Unable to find any table headers in the Pickup List")
        
        if not header_names:
            # Debug: Print page source to see what's actually there
            logging.error("âŒ No headers found. Current page source:")
            logging.error(driver.page_source[:2000])  # First 2000 chars
            raise RuntimeError("Unable to find any table headers in the Pickup List")
        # Map header text â†’ 1-based column index
        col_idx = {name: idx+1 for idx, name in enumerate(header_names)}
        # Verify required headers are present
        for needed in ("Account", "Close", "Address Line 1"):
            if needed not in col_idx:
                raise RuntimeError(
                    f"Missing expected header '{needed}'; found: {header_names}"
                )
        # Assigned Route sits immediately after the Close column
        asg_idx = col_idx["Close"] + 1

        # 5) Grab the number of regulars from the "Regulars" badge with multiple fallback methods
        reg_count = None
        timeout = time.time() + 30
        
        # Method 1: Try the specific space8 selector
        try:
            span = driver.find_element(By.CSS_SELECTOR, "span[id='space8']")
            text = span.text.strip()
            logging.info(f"Found space8 span text: '{text}'")
            if text.isdigit():
                reg_count = int(text)
                logging.info(f"âœ… Found regulars count (method 1): {reg_count}")
            else:
                logging.warning(f"âš ï¸ space8 span contains non-numeric text: '{text}'")
        except Exception as e:
            logging.warning(f"âš ï¸ Method 1 failed: {e}")
        
        # Method 2: Look for any span containing a number that might be the regulars count
        if reg_count is None:
            try:
                spans = driver.find_elements(By.TAG_NAME, "span")
                for span in spans:
                    text = span.text.strip()
                    if text.isdigit() and len(text) <= 4:  # Reasonable count range
                        reg_count = int(text)
                        logging.info(f"âœ… Found regulars count (method 2): {reg_count}")
                        break
            except Exception as e:
                logging.warning(f"âš ï¸ Method 2 failed: {e}")
        
        # Method 3: Look for text containing "Regular" and extract number
        if reg_count is None:
            try:
                elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Regular')]")
                for element in elements:
                    text = element.text.strip()
                    # Look for patterns like "Regular (123)" or "123 Regular"
                    import re
                    match = re.search(r'(\d+)', text)
                    if match:
                        reg_count = int(match.group(1))
                        logging.info(f"âœ… Found regulars count (method 3): {reg_count}")
                        break
            except Exception as e:
                logging.warning(f"âš ï¸ Method 3 failed: {e}")
        
        # Method 4: Count regular pickup rows in the table
        if reg_count is None:
            try:
                # Look for rows with Disp# starting with 6 (regular pickups)
                regular_rows = driver.find_elements(
                    By.XPATH, 
                    "//tr[.//div[starts-with(text(), '6') and string-length(text()) >= 4]]"
                )
                if regular_rows:
                    reg_count = len(regular_rows)
                    logging.info(f"âœ… Found regulars count (method 4): {reg_count} by counting rows")
            except Exception as e:
                logging.warning(f"âš ï¸ Method 4 failed: {e}")
        
        # Method 5: Manual fallback - assume a reasonable number
        if reg_count is None:
            logging.warning("âš ï¸ Could not determine regulars count automatically, using fallback value")
            reg_count = 50  # Reasonable fallback
            logging.info(f"âœ… Using fallback regulars count: {reg_count}")
        
        if reg_count is None:
            raise RuntimeError("Unable to determine regulars count with any method")

        # 6+7) Compare DSM CSV data with SIMS data
        mismatches = []
        no_matches = []
        processed_disp_numbers = []  # Track all dispatch numbers processed
        
        # Check if we have DSM CSV data
        if hasattr(self, 'dsm_df') and self.dsm_df is not None:
            logging.info("âœ… Using DSM CSV data for comparison")
            print(f"ðŸ” Starting CSV comparison: DSM rows={len(self.dsm_df)}, SIMS rows={len(self.sims_df)}", flush=True)
            
            try:
                # Process DSM CSV data
                for index, dsm_row in self.dsm_df.iterrows():
                    # Check if stop was requested
                    if self.sort_stop_event.is_set():
                        print("ðŸ›‘ Auto-sort process stopped by user", flush=True)
                        break
                    
                    # Get DSM values from CSV row
                    try:
                        # Extract values from DSM CSV row (adjust column names as needed)
                        # Extract dispatch number from column C ("Disp #") - this is the specific column mentioned by user
                        dsm_disp_no = str(dsm_row.get('Disp #', dsm_row.get('Disp#', dsm_row.get('Dispatch Number', dsm_row.get('Dispatch', dsm_row.get('Disp', ''))))))
                        
                        # Debug: Print all column names and values for first few rows to identify dispatch column
                        if index < 3:
                            print(f"DEBUG Row {index + 1} - All columns and values:")
                            for col_name, col_value in dsm_row.items():
                                print(f"  {col_name}: {col_value}")
                            print("---")
                        
                        # Clean up the dispatch number - extract just the 4-digit number
                        if dsm_disp_no and dsm_disp_no != 'nan':
                            # Remove any non-digit characters and take first 4 digits
                            import re
                            digits_only = re.sub(r'[^\d]', '', dsm_disp_no)
                            if len(digits_only) >= 4:
                                dsm_disp_no = digits_only[:4]
                            elif len(digits_only) > 0:
                                dsm_disp_no = digits_only.zfill(4)  # Pad with zeros if less than 4 digits
                            else:
                                dsm_disp_no = "0000"  # Default if no digits found
                        else:
                            dsm_disp_no = "0000"  # Default for empty/nan values
                        dsm_account = str(dsm_row.get('Account', dsm_row.get('Acc#/Loc Id', '')))
                        dsm_customer = str(dsm_row.get('Customer', dsm_row.get('Co Name', '')))
                        dsm_address = str(dsm_row.get('Address Line 1', dsm_row.get('Street', '')))
                        dsm_ready = str(dsm_row.get('Ready', dsm_row.get('Rdy Tm', '')))
                        dsm_close = str(dsm_row.get('Close', dsm_row.get('Cls Tm', '')))
                        dsm_route = str(dsm_row.get('Asg Route', dsm_row.get('Primary Route', '')))
                        dsm_status = str(dsm_row.get('Status', 'Active')).lower()
                        
                        # Log all dispatch numbers being processed (including skipped ones)
                        print(f"\n--- Processing DSM CSV row {index + 1}: Disp# {dsm_disp_no} ---", flush=True)
                        
                        # Track this dispatch number
                        processed_disp_numbers.append(dsm_disp_no)
                        
                        # Skip if not active (including active-updated)
                        if dsm_status not in ['active', 'active-updated']:
                            print(f"  â­ï¸ Skipping Disp# {dsm_disp_no} - Status: '{dsm_status}' (not active)", flush=True)
                            continue
                        print(f"DSM values: Account={dsm_account}, Customer={dsm_customer}, Route={dsm_route}", flush=True)
                        
                        # Update progress bar
                        progress_percentage = ((index + 1) / len(self.dsm_df)) * 100
                        if progress_percentage > 100:
                            print(f"  âš ï¸ Progress capped at 100% (row {index + 1})", flush=True)
                                                
                        # Log progress every 50 rows for monitoring
                        if (index + 1) % 50 == 0:
                            print(f"  ðŸ“Š Progress: Row {index + 1}/{len(self.dsm_df)} ({progress_percentage:.1f}% complete)", flush=True)
                            
                    except Exception as e:
                        print(f"  âš ï¸ Error extracting DSM values for row {index + 1}: {e}", flush=True)
                        continue
                    
                    # Enhanced matching logic with multiple fallback strategies
                    df_match = None
                    match_method = ""
                    active_dispatch_row = None  # Initialize active dispatch row variable
                    
                    # Strategy 1: Perfect match by Account + Address + Times
                    if dsm_account:
                        df_match = self.sims_df[
                            (self.sims_df[column_mapping['Account']] == dsm_account) &
                            (self.sims_df[column_mapping['Street']] == dsm_address) &
                            (self.sims_df[column_mapping['Ready']].astype(str).str.replace('.0', '') == dsm_ready) &
                            (self.sims_df[column_mapping['Close']].astype(str).str.replace('.0', '') == dsm_close)
                        ]
                        if not df_match.empty:
                            match_method = "Account + Address + Times"
                
                    # Strategy 2: Account + Address (if times don't match)
                    if df_match is None or df_match.empty:
                        if dsm_account:
                            df_match = self.sims_df[
                                (self.sims_df[column_mapping['Account']] == dsm_account) &
                                (self.sims_df[column_mapping['Street']] == dsm_address)
                            ]
                            if not df_match.empty:
                                match_method = "Account + Address"
                
                    # Strategy 3: Account + Customer + Times
                    if df_match is None or df_match.empty:
                        if dsm_account:
                            df_match = self.sims_df[
                                (self.sims_df[column_mapping['Account']] == dsm_account) &
                                (self.sims_df[column_mapping['Customer']] == dsm_customer) &
                                (self.sims_df[column_mapping['Ready']].astype(str).str.replace('.0', '') == dsm_ready) &
                                (self.sims_df[column_mapping['Close']].astype(str).str.replace('.0', '') == dsm_close)
                            ]
                            if not df_match.empty:
                                match_method = "Account + Customer + Times"
                
                                        # Strategy 4: Address + Customer + Times (for cases with valid account)
                    if df_match is None or df_match.empty:
                        if dsm_account and dsm_account != 'nan':
                            df_match = self.sims_df[
                                (self.sims_df[column_mapping['Street']] == dsm_address) &
                                (self.sims_df[column_mapping['Customer']] == dsm_customer) &
                                (self.sims_df[column_mapping['Ready']].astype(str).str.replace('.0', '') == dsm_ready) &
                                (self.sims_df[column_mapping['Close']].astype(str).str.replace('.0', '') == dsm_close)
                            ]
                            if not df_match.empty:
                                match_method = "Address + Customer + Times"

                    # Strategy 5: Address + Times (for cases with valid account)
                    if df_match is None or df_match.empty:
                        if dsm_account and dsm_account != 'nan':
                            df_match = self.sims_df[
                                (self.sims_df[column_mapping['Street']] == dsm_address) &
                                (self.sims_df[column_mapping['Ready']].astype(str).str.replace('.0', '') == dsm_ready) &
                                (self.sims_df[column_mapping['Close']].astype(str).str.replace('.0', '') == dsm_close)
                            ]
                            if not df_match.empty:
                                match_method = "Address + Times"
                    
                    # Strategy 6: Address only (for cases with no account number - will validate times later)
                    if df_match is None or df_match.empty:
                        if not dsm_account or dsm_account == 'nan':
                            df_match = self.sims_df[self.sims_df[column_mapping['Street']] == dsm_address]
                            if not df_match.empty:
                                match_method = "Address only (no account)"
                
                                        # Strategy 7: Account only (last resort for cases with valid account)
                    if df_match is None or df_match.empty:
                        if dsm_account and dsm_account != 'nan':
                            df_match = self.sims_df[self.sims_df[column_mapping['Account']] == dsm_account]
                            if not df_match.empty:
                                match_method = "Account only"

                    # Strategy 8: Address only (final fallback for any remaining cases)
                    if df_match is None or df_match.empty:
                        df_match = self.sims_df[self.sims_df[column_mapping['Street']] == dsm_address]
                        if not df_match.empty:
                            match_method = "Address only (fallback)"
                
                    # Check if we found a match
                    if df_match is None or df_match.empty:
                        print(f"  âŒ NO MATCH FOUND in SIMS data for Disp# {dsm_disp_no}", flush=True)
                        print(f"     DSM values: Account={dsm_account}, Address='{dsm_address}', Customer='{dsm_customer}', Ready={dsm_ready}, Close={dsm_close}", flush=True)
                        
                        # Add no-match to results
                        no_match_text = f"âŒ NO MATCH: Disp#{dsm_disp_no} | Route: {dsm_route} | Customer: {dsm_customer}"
                        self.add_result(no_match_text)
                        no_matches.append(dsm_disp_no)
                        continue
                
                    # If multiple matches found, try to find the best match based on times
                    if len(df_match) > 1:
                        print(f"  âš ï¸ Multiple SIMS matches found ({len(df_match)}), checking for best time match", flush=True)
                        
                        # For address-only matches (no account), try to find a match with matching times
                        if match_method == "Address only (no account)":
                            best_match = None
                            best_match_score = -1  # -1 = no match, 0 = partial match, 1 = perfect match
                            
                            for idx, row in df_match.iterrows():
                                # Get times from this potential match
                                match_ready_raw = row[column_mapping['Ready']]
                                match_close_raw = row[column_mapping['Close']]
                                
                                # Normalize times
                                def norm_time(val):
                                    s = str(val).strip()
                                    try:
                                        f = float(s)
                                        if f.is_integer():
                                            return str(int(f))
                                    except:
                                        pass
                                    return s
                                
                                match_ready = norm_time(match_ready_raw)
                                match_close = norm_time(match_close_raw)
                                
                                route_val = row[column_mapping['Asg Route']]
                                try:
                                    nr = float(route_val)
                                    route_str = str(int(nr)) if nr.is_integer() else str(nr)
                                except Exception:
                                    route_str = str(route_val).strip()
                                
                                # Calculate match score with TIME priority for address-only matches
                                score = 0
                                
                                # Time matching (0.3 points each, max 0.6) - HIGHEST PRIORITY for address-only
                                if match_ready == dsm_ready:
                                    score += 0.3
                                if match_close == dsm_close:
                                    score += 0.3
                                
                                # Route matching (0.3 points) - SECONDARY priority for address-only
                                normalized_match_route = str(int(float(route_str))) if route_str.replace('.','').isdigit() else route_str
                                normalized_dsm_route_temp = str(int(float(dsm_route))) if dsm_route.replace('.','').isdigit() else dsm_route
                                if normalized_match_route == normalized_dsm_route_temp:
                                    score += 0.3
                                
                                # Customer matching (0.1 points) - LOWEST priority for address-only
                                customer_val = row[column_mapping['Customer']]
                                customer_str = str(customer_val).strip()
                                if customer_str == dsm_customer:
                                    score += 0.1
                                
                                print(f"    ðŸ” Match option: Route={route_str}, Times={match_ready}-{match_close}, Customer={customer_str}, Score={score:.1f}", flush=True)
                                
                                # Perfect match (times + route + customer match)
                                if score == 1.0:  # 0.6 (times) + 0.3 (route) + 0.1 (customer) = 1.0
                                    best_match = row
                                    best_match_score = score
                                    print(f"    âœ… Perfect match found! Using Route={route_str}", flush=True)
                                    break
                                # Better match (higher score)
                                elif score > best_match_score:
                                    best_match = row
                                    best_match_score = score
                                # Same score - prioritize time matches over route matches for address-only
                                elif score == best_match_score and score >= 0.6:  # Has time match
                                    # Check if current best match has time match
                                    if best_match is not None:
                                        current_best_ready_raw = best_match[column_mapping['Ready']]
                                        current_best_close_raw = best_match[column_mapping['Close']]
                                        current_best_ready = norm_time(current_best_ready_raw)
                                        current_best_close = norm_time(current_best_close_raw)
                                        
                                        current_best_times_match = (current_best_ready == dsm_ready) and (current_best_close == dsm_close)
                                        current_match_times_match = (match_ready == dsm_ready) and (match_close == dsm_close)
                                        
                                        # If current best doesn't have time match but new match does, replace it
                                        if not current_best_times_match and current_match_times_match:
                                            best_match = row
                                            best_match_score = score
                                            print(f"    ðŸ”„ Replacing with time match: Times={match_ready}-{match_close}", flush=True)
                            
                            # Use the best match found, or first if no matches
                            if best_match is not None:
                                df_match = best_match.to_frame().T
                                if best_match_score == 1.0:
                                    print(f"  âœ… Selected perfect match (times + route + customer)", flush=True)
                                elif best_match_score >= 0.6:
                                    print(f"  âœ… Selected time match (score: {best_match_score:.1f})", flush=True)
                                elif best_match_score >= 0.3:
                                    print(f"  ðŸ“‹ Selected route match (score: {best_match_score:.1f})", flush=True)
                                elif best_match_score > 0:
                                    print(f"  ðŸ“‹ Selected partial match (score: {best_match_score:.1f})", flush=True)
                                else:
                                    print(f"  âš ï¸ No good matches found, using first available", flush=True)
                            else:
                                print(f"  âš ï¸ No suitable match found, using first match", flush=True)
                                df_match = df_match.head(1)
                        else:
                            # For other match methods, just use the first match
                            print(f"  ðŸ“‹ Using first match for {match_method}", flush=True)
                            df_match = df_match.head(1)
                
                    print(f"  âœ“ Matched via {match_method}", flush=True)

                    # Retrieve and normalize the Excel Primary Route
                    # Use active dispatch row if available, otherwise use the matched row
                    route_row = active_dispatch_row if active_dispatch_row is not None else df_match.iloc[0]
                    
                    raw_route = route_row[column_mapping['Asg Route']]
                    try:
                        nr = float(raw_route)
                        excel_route = str(int(nr)) if nr.is_integer() else str(nr)
                    except Exception:
                        excel_route = str(raw_route).strip()
                    
                    if active_dispatch_row is not None:
                        print(f"Excel values: PrimaryRoute = {excel_route} (from active dispatch)", flush=True)
                    else:
                        print(f"Excel values: PrimaryRoute = {excel_route}", flush=True)

                    # â”€â”€â”€ Compare DSM vs SIMS data â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    excel_street = str(df_match[column_mapping['Street']].iloc[0]).strip()
                    print(f"DSM values: Address = '{dsm_address}'", flush=True)
                    print(f"SIMS values: Street = '{excel_street}'", flush=True)
                    print(f"DSM times: Ready={dsm_ready}, Close={dsm_close}", flush=True)

                    # â”€â”€â”€ Pull Excel times â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    excel_ready_raw = df_match[column_mapping['Ready']].iloc[0]
                    excel_close_raw = df_match[column_mapping['Close']].iloc[0]
                    # normalize Excel times (drop trailing .0 if integer)
                    def norm(val):
                        s = str(val).strip()
                        try:
                            f = float(s)
                            if f.is_integer():
                                return str(int(f))
                        except:
                            pass
                        return s
                    excel_ready = norm(excel_ready_raw)
                    excel_close = norm(excel_close_raw)
                    print(f"Excel times: Rdy Tm={excel_ready}, Cls Tm={excel_close}", flush=True)

                    # â”€â”€â”€ Compare ONLY the route (SIMS is the source of truth) â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    # Normalize route numbers by removing leading zeros for comparison
                    def normalize_route(route_str):
                        if not route_str:
                            return ""
                        try:
                            # Convert to int to remove leading zeros, then back to string
                            return str(int(route_str))
                        except:
                            # If conversion fails, return original string
                            return str(route_str).strip()
                
                    normalized_excel_route = normalize_route(excel_route)
                    normalized_dsm_route = normalize_route(dsm_route)
                    
                    route_mismatch = False
                    if normalized_excel_route != normalized_dsm_route:
                        route_mismatch = True
                        print(f"  âŒ ROUTE MISMATCH! DSM={dsm_route} vs SIMS={excel_route} (normalized: {normalized_dsm_route} vs {normalized_excel_route})", flush=True)
                        
                        # Don't create mismatch record yet - wait to see if cross-reference resolves it
                        # We'll create it later if no cross-reference is found or if cross-reference also fails
                    else:
                        print(f"  âœ… ROUTE MATCHES! Route={dsm_route}", flush=True)
                    
                    # â”€â”€â”€ Additional time validation for address-only matches (no account) â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if match_method == "Address only (no account)":
                        # For cases with no account number, validate times separately
                        times_match = (dsm_ready == excel_ready) and (dsm_close == excel_close)
                        if not times_match:
                            print(f"  âš ï¸ TIME MISMATCH! DSM times: Ready={dsm_ready}, Close={dsm_close} vs SIMS times: Ready={excel_ready}, Close={excel_close}", flush=True)
                            print(f"  ðŸ“‹ Note: Matched by address only (no account), but times differ", flush=True)
                        else:
                            print(f"  âœ… TIMES MATCH! Ready={dsm_ready}, Close={dsm_close}", flush=True)
                
                    # â”€â”€â”€ Check day of week columns for current day â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    try:
                        # Get current day of week (0=Monday, 6=Sunday)
                        from datetime import datetime
                        current_day_num = datetime.now().weekday()  # 0=Monday, 1=Tuesday, etc.
                        day_abbreviations = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
                        current_day_abbr = day_abbreviations[current_day_num]
                        
                        # Debug: Print all SIMS columns to help identify day columns
                        if index < 3:  # Only print for first few rows to avoid spam
                            print(f"  ðŸ” All SIMS columns for debugging: {list(self.sims_df.columns)}", flush=True)
                        
                        # Try multiple methods to find the current day's column
                        current_day_value = None
                        current_day_column = None
                        current_day_str = ""
                        
                        # Method 1: Try the hardcoded column names
                        day_column_mapping = {
                            0: "Unnamed: 22",  # Monday
                            1: "Unnamed: 23",  # Tuesday  
                            2: "Unnamed: 24",  # Wednesday
                            3: "Unnamed: 25",  # Thursday
                            4: "Unnamed: 26",  # Friday
                            5: "Unnamed: 27",  # Saturday
                            6: "Unnamed: 28",  # Sunday
                        }
                        
                        if current_day_num in day_column_mapping:
                            column_name = day_column_mapping[current_day_num]
                            if column_name in self.sims_df.columns:
                                current_day_column = column_name
                                current_day_value = df_match[current_day_column].iloc[0]
                                current_day_str = str(current_day_value).strip()
                                print(f"  âœ… Found day column using hardcoded mapping: {current_day_column}", flush=True)
                        
                        # Method 2: Search for columns containing day abbreviations
                        if current_day_value is None:
                            day_patterns = {
                                0: ["mon", "monday", "mo"],
                                1: ["tue", "tuesday", "tu"], 
                                2: ["wed", "wednesday", "we"],
                                3: ["thu", "thursday", "th"],
                                4: ["fri", "friday", "fr"],
                                5: ["sat", "saturday", "sa"],
                                6: ["sun", "sunday", "su"]
                            }
                            
                            if current_day_num in day_patterns:
                                patterns = day_patterns[current_day_num]
                                for col in self.sims_df.columns:
                                    col_lower = str(col).lower()
                                    if any(pattern in col_lower for pattern in patterns):
                                        current_day_column = col
                                        current_day_value = df_match[current_day_column].iloc[0]
                                        current_day_str = str(current_day_value).strip()
                                        print(f"  âœ… Found day column using pattern matching: {current_day_column}", flush=True)
                                        break
                        
                        # Method 3: Look for columns with day abbreviations in the header
                        if current_day_value is None:
                            # Try to find columns that might contain the day abbreviation
                            for col in self.sims_df.columns:
                                col_str = str(col).strip()
                                if col_str == current_day_abbr or col_str.lower() == current_day_abbr.lower():
                                    current_day_column = col
                                    current_day_value = df_match[current_day_column].iloc[0]
                                    current_day_str = str(current_day_value).strip()
                                    print(f"  âœ… Found day column using abbreviation match: {current_day_column}", flush=True)
                                    break
                        
                        # Method 4: Look for columns with day names
                        if current_day_value is None:
                            day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                            if current_day_num < len(day_names):
                                day_name = day_names[current_day_num]
                                for col in self.sims_df.columns:
                                    col_str = str(col).strip()
                                    if day_name.lower() in col_str.lower():
                                        current_day_column = col
                                        current_day_value = df_match[current_day_column].iloc[0]
                                        current_day_str = str(current_day_value).strip()
                                        print(f"  âœ… Found day column using day name: {current_day_column}", flush=True)
                                        break
                        
                        # Debug: Print current day info
                        print(f"  ðŸ” Current day: {current_day_abbr} (day #{current_day_num})", flush=True)
                        if current_day_column:
                            print(f"  ðŸ” Found column: {current_day_column}", flush=True)
                            print(f"  ðŸ” Column value: '{current_day_str}' (type: {type(current_day_value)})", flush=True)
                            
                            # For first few rows, show all day column values to understand the data format
                            if index < 3:
                                print(f"  ðŸ” DEBUG: All day column values for this pickup:", flush=True)
                                for day_num, day_name in enumerate(['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']):
                                    if day_num in day_column_mapping:
                                        col_name = day_column_mapping[day_num]
                                        if col_name in self.sims_df.columns:
                                            day_value = df_match[col_name].iloc[0]
                                            day_str = str(day_value).strip()
                                            print(f"    {day_name}: '{day_str}' (raw: {day_value}, type: {type(day_value)})", flush=True)
                                
                                # Also show what the current day's value should indicate
                                print(f"  ðŸ” DEBUG: Current day ({current_day_abbr}) value analysis:", flush=True)
                                print(f"    - Raw value: {current_day_value} (type: {type(current_day_value)})", flush=True)
                                print(f"    - String value: '{current_day_str}'", flush=True)
                                print(f"    - Is float nan: {isinstance(current_day_value, float) and math.isnan(current_day_value) if current_day_value is not None else 'N/A'}", flush=True)
                                print(f"    - Is active: {is_active_today}", flush=True)
                        else:
                            print(f"  âš ï¸ Could not find column for {current_day_abbr}", flush=True)
                            print(f"  ðŸ” Available columns: {list(self.sims_df.columns)}", flush=True)
                        
                        # Check if pickup is active on current day (non-empty value means active)
                        # Handle pandas nan values (which are floats, not None)
                        import math
                        
                        # More robust check for active status
                        is_active_today = False
                        if current_day_value is not None:
                            # Check if it's a pandas nan (float)
                            if isinstance(current_day_value, float) and math.isnan(current_day_value):
                                is_active_today = False
                            elif current_day_str and current_day_str.lower() not in ['nan', 'none', '', '0', 'false', 'no']:
                                is_active_today = True
                            else:
                                is_active_today = False
                        else:
                            is_active_today = False
                        
                        # If current dispatch is not active today, check other dispatch numbers with same account
                        if not is_active_today and dsm_account:
                            print(f"  ðŸ” Current dispatch not active today, checking other dispatch numbers with account {dsm_account}...", flush=True)
                            
                            # Normalize account numbers by removing decimal part for comparison
                            def normalize_account(account_str):
                                if not account_str or account_str == 'nan':
                                    return ''
                                # Remove .0 suffix if present
                                account_str = str(account_str).strip()
                                if account_str.endswith('.0'):
                                    return account_str[:-2]
                                return account_str
                            
                            normalized_current_account = normalize_account(dsm_account)
                            print(f"  ðŸ” Normalized current account: '{normalized_current_account}'", flush=True)
                            
                            # Find all other dispatch numbers with the same account in SIMS data
                            other_dispatches = []
                            for other_index, other_row in self.sims_df.iterrows():
                                other_account = str(other_row.get(column_mapping['Account'], '')).strip()
                                normalized_other_account = normalize_account(other_account)
                                
                                if normalized_other_account == normalized_current_account and normalized_current_account:
                                    # Get the dispatch identifier - try multiple approaches
                                    other_disp_identifier = "Unknown"
                                    
                                    # Method 1: Try the first column which often contains dispatch info
                                    first_col_value = str(other_row.iloc[0]).strip()
                                    if first_col_value and first_col_value != 'nan':
                                        # Extract dispatch number from first column (e.g., "369055941 DIGITAL T18155 NW MIAMI")
                                        import re
                                        match = re.search(r'(\d{9})', first_col_value)  # Look for 9-digit dispatch number
                                        if match:
                                            other_disp_identifier = match.group(1)
                                    
                                    # Method 2: Try specific dispatch columns if Method 1 failed
                                    if other_disp_identifier == "Unknown":
                                        for col in self.sims_df.columns:
                                            col_str = str(col).lower()
                                            if 'disp' in col_str or 'dispatch' in col_str:
                                                col_value = str(other_row.get(col, '')).strip()
                                                if col_value and col_value != 'nan':
                                                    other_disp_identifier = col_value
                                                    break
                                    
                                    # Method 3: Use row index as identifier if all else fails
                                    if other_disp_identifier == "Unknown":
                                        other_disp_identifier = f"Row_{other_index + 1}"
                                    
                                    other_dispatches.append((other_index, other_disp_identifier))
                            
                            if other_dispatches:
                                print(f"  ðŸ” Found {len(other_dispatches)} other dispatch numbers with same account: {[disp[1] for disp in other_dispatches]}", flush=True)
                                
                                # Check each other dispatch for day of week activity
                                for other_index, other_disp_identifier in other_dispatches:
                                    print(f"  ðŸ” Checking other dispatch '{other_disp_identifier}' for day activity...", flush=True)
                                    
                                    # Get the other dispatch row directly from SIMS data
                                    other_row = self.sims_df.iloc[other_index]
                                    
                                    # Check day of week for this other dispatch
                                    other_day_value = None
                                    if current_day_column and current_day_column in self.sims_df.columns:
                                        other_day_value = other_row[current_day_column]
                                        other_day_str = str(other_day_value).strip()
                                        
                                        # Check if this other dispatch is active today
                                        other_is_active = False
                                        if other_day_value is not None:
                                            if isinstance(other_day_value, float) and math.isnan(other_day_value):
                                                other_is_active = False
                                            elif other_day_str and other_day_str.lower() not in ['nan', 'none', '', '0', 'false', 'no']:
                                                other_is_active = True
                                        
                                        if other_is_active:
                                            print(f"  âœ… Found active dispatch '{other_disp_identifier}' with same account - using this for route comparison", flush=True)
                                            is_active_today = True
                                            current_day_str = other_day_str
                                            current_day_value = other_day_value
                                            active_dispatch_row = other_row  # Store the active dispatch row for route comparison
                                            
                                            # Re-match using the active dispatch's data
                                            print(f"  ðŸ”„ Re-matching using active dispatch data...", flush=True)
                                            
                                            # Get the active dispatch's data for matching
                                            active_account = str(active_dispatch_row.get(column_mapping['Account'], '')).strip()
                                            active_customer = str(active_dispatch_row.get(column_mapping['Customer'], '')).strip()
                                            active_address = str(active_dispatch_row.get(column_mapping['Street'], '')).strip()
                                            active_ready = str(active_dispatch_row.get(column_mapping['Ready'], '')).strip()
                                            active_close = str(active_dispatch_row.get(column_mapping['Close'], '')).strip()
                                            
                                            # Use the active dispatch's route for comparison
                                            print(f"  âœ… Using active dispatch to confirm pickup is active today", flush=True)
                                            print(f"  ðŸ“‹ Will compare active dispatch route with pickup list route", flush=True)
                                            
                                            # Get the active dispatch's route from SIMS data
                                            active_route = active_dispatch_row.get(column_mapping['Asg Route'], None)
                                            if active_route is not None:
                                                try:
                                                    nr = float(active_route)
                                                    active_route_str = str(int(nr)) if nr.is_integer() else str(nr)
                                                except Exception:
                                                    active_route_str = str(active_route).strip()
                                                
                                                print(f"  ðŸ”„ Active dispatch route: {active_route_str}", flush=True)
                                                print(f"  ðŸ”„ Pickup list route: {dsm_route}", flush=True)
                                                
                                                # Compare active dispatch route with pickup list route
                                                normalized_active_route = normalize_route(active_route_str)
                                                normalized_dsm_route = normalize_route(dsm_route)
                                                
                                                if normalized_active_route != normalized_dsm_route:
                                                    print(f"  âŒ ROUTE MISMATCH! Pickup List={dsm_route} vs Active Dispatch={active_route_str} (normalized: {normalized_dsm_route} vs {normalized_active_route})", flush=True)
                                                    
                                                    # Create mismatch record using active dispatch route
                                                    mismatch_record = {
                                                        'disp_no': dsm_disp_no,
                                                        'account': dsm_account,
                                                        'address': dsm_address,
                                                        'customer': dsm_customer,
                                                        'dsm_route': dsm_route,
                                                        'excel_route': active_route_str,
                                                        'ready_time': dsm_ready,
                                                        'close_time': dsm_close,
                                                        'match_method': f"Active Dispatch - {match_method}"
                                                    }
                                                    
                                                    mismatches.append(mismatch_record)
                                                    
                                                    # Add mismatch to results (simplified format with Disp#)
                                                    mismatch_text = f"âŒ Disp#{dsm_disp_no} | Route: {dsm_route} â†’ {active_route_str} | Customer: {dsm_customer}"
                                                    self.add_result(mismatch_text)
                                                else:
                                                    print(f"  âœ… ROUTE MATCHES! Route={dsm_route} (using active dispatch)", flush=True)
                                                    
                                                    # Remove any existing mismatch for this dispatch number since routes now match
                                                    mismatches = [m for m in mismatches if m['disp_no'] != dsm_disp_no]
                                                    
                                                    # Remove the mismatch result from the UI as well
                                                    # Find and remove the mismatch result frame
                                                    for i, result in enumerate(self.result_frames):
                                                        try:
                                                            if result['text'] and f"Disp#{dsm_disp_no}" in result['text'] and "âŒ" in result['text']:
                                                                result['frame'].destroy()
                                                                self.result_frames.pop(i)
                                                                break
                                                        except Exception:
                                                            # Frame might already be destroyed, skip it
                                                            continue
                                                
                                                # Skip the original route comparison since we've already handled it
                                                route_mismatch = False
                                                continue
                                            else:
                                                print(f"  âš ï¸ Could not get route from active dispatch", flush=True)
                                            
                                            break
                                        else:
                                            print(f"  âŒ Other dispatch '{other_disp_identifier}' also not active today", flush=True)
                                    else:
                                        print(f"  âš ï¸ Could not check day column for other dispatch '{other_disp_identifier}'", flush=True)
                            else:
                                print(f"  â„¹ï¸ No other active dispatch numbers found with same account", flush=True)
                        
                        if current_day_value is not None:
                            if is_active_today:
                                print(f"  ðŸ“… PICKUP ACTIVE TODAY ({current_day_abbr}) - Value: '{current_day_str}' (raw: {current_day_value})", flush=True)
                            else:
                                print(f"  ðŸ“… PICKUP NOT ACTIVE TODAY ({current_day_abbr}) - Value: '{current_day_str}' (raw: {current_day_value})", flush=True)
                        else:
                            print(f"  âš ï¸ Could not determine if pickup is active today - no day column found", flush=True)
                        
                        # If we had a route mismatch, create the mismatch record now
                        if route_mismatch:
                            # Create mismatch record for the original route comparison
                            mismatch_record = {
                                'disp_no': dsm_disp_no,
                                'account': dsm_account,
                                'address': dsm_address,
                                'customer': dsm_customer,
                                'dsm_route': dsm_route,
                                'excel_route': excel_route,
                                'ready_time': dsm_ready,
                                'close_time': dsm_close,
                                'match_method': match_method
                            }
                            
                            mismatches.append(mismatch_record)
                            
                            # Add mismatch to results (simplified format with Disp#)
                            mismatch_text = f"âŒ Disp#{dsm_disp_no} | Route: {dsm_route} â†’ {excel_route} | Customer: {dsm_customer}"
                            self.add_result(mismatch_text)
                            
                    except Exception as e:
                        print(f"  âš ï¸ Could not check day of week: {e}", flush=True)
                        import traceback
                        traceback.print_exc()
                
                    # Update the progress bar based on current row being processed
                    progress_percentage = ((index + 1) / len(self.dsm_df)) * 100
                    
                    # Debug: Print progress info
                    print(f"  ðŸ” Progress Debug: Row {index + 1}/{len(self.dsm_df)}, progress={progress_percentage:.1f}%", flush=True)
                    
                    # Ensure progress doesn't exceed 100%
                    if progress_percentage > 100:
                        progress_percentage = 100
                        print(f"  âš ï¸ Progress capped at 100% (Row {index + 1})", flush=True)
                    
                    # Update custom progress bar more frequently for smoother display
                    # Round to nearest 1% for smoother updates
                    display_percentage = round(progress_percentage, 1)
                    self.update_custom_progress(display_percentage)
                    self.update_idletasks()
                    
                    # Performance optimization: Force garbage collection every 200 rows
                    if (index + 1) % 200 == 0:
                        try:
                            import gc
                            gc.collect()
                            print(f"  ðŸ—‘ï¸ Memory cleanup at row {index + 1}", flush=True)
                        except Exception:
                            pass
                        
            except Exception as e:
                print(f"Error during CSV processing: {e}")
                import traceback
                traceback.print_exc()
            finally:
                # final summary
                print("\n=== Regular Pickup Check Complete ===", flush=True)
                
                # Ensure progress bar reaches 100% at completion
                self.update_custom_progress(100.0)
                self.update_idletasks()
            
            # Only show results if there are mismatches
            if mismatches:
                # Add final summary to results
                self.add_result("")
                self.add_result("=" * 60)
                self.add_result("ðŸ“Š FINAL SUMMARY")
                self.add_result("=" * 60)
                
                summary_text = f"Total stops mismatched/out of place: {len(mismatches)}"
                self.add_result(summary_text)
                print(f"\nðŸ“Š FINAL SUMMARY:", flush=True)
                print(f"Total route mismatches found: {len(mismatches)}", flush=True)
                
                if no_matches:
                    no_match_summary = f"Total no-matches found: {len(no_matches)}"
                    self.add_result(no_match_summary)
                    print(f"Total no-matches found: {len(no_matches)}", flush=True)
            else:
                # No mismatches found - show success message on UI
                self.add_result("")
                self.add_result("=" * 60)
                self.add_result("âœ… SUCCESS")
                self.add_result("=" * 60)
                self.add_result("All regular pickup routes match the SIMS report.")
                print("âœ… All regular pickup routes match the SIMS report.", flush=True)

        # TODO: now continue your per-route sorting, updating self.sort_progress.set(percent) as you go
        
        # Update UI based on whether process was stopped or completed
        if self.sort_stop_event.is_set():
            self.sort_status_label.configure(text="Stopped", text_color="red")
        else:
            self.sort_status_label.configure(text="Complete", text_color=BRAND_ORANGE)
        
        # Reset button states
        self.sort_start_button.configure(state="normal")
        self.sort_stop_button.configure(state="disabled")
        
        # Cleanup
        try:
            if hasattr(self, 'driver') and self.driver:
                self.driver.quit()
                self.driver = None
        except:
            pass

    # â”€â”€â”€ Stop Movement Automation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    
    def start_stop_movement(self):
        """Start the stop movement automation process"""
        from tkinter import messagebox
        
        # Check if user is logged in
        if not self.username or not self.password:
            messagebox.showerror("Not Logged In", "Please log in first before using stop movement automation.")
            return

        # Check if we have mismatch data to process
        if not hasattr(self, 'result_frames') or not self.result_frames:
            messagebox.showerror("No Data", "Please run the route comparison first to get mismatch data.")
            return

        # Get mismatch data from results (only selected checkboxes)
        mismatch_data = []
        selected_count = 0
        total_mismatches = 0
        
        for result in self.result_frames:
            text = result['text']
            if text and text.strip().startswith("âŒ") and "Disp#" in text:
                total_mismatches += 1
                
                # Check if this mismatch is selected via checkbox
                if result['checkbox'] is not None and result['checkbox_var'].get():
                    selected_count += 1
                    
                    # Parse the mismatch data
                    # Format: "âŒ Disp#6192 | Route: 459 â†’ 779 | Customer: CLOSED SOUTHWEST FLORIDA LLC"
                    try:
                        # Extract dispatch number
                        disp_match = re.search(r'Disp#(\d+)', text)
                        if disp_match:
                            disp_no = disp_match.group(1)
                            
                            # Extract route information
                            route_match = re.search(r'Route: (\d+) â†’ (\d+)', text)
                            if route_match:
                                current_route = route_match.group(1)
                                correct_route = route_match.group(2)
                                
                                # Extract customer name
                                customer_match = re.search(r'Customer: (.+)$', text)
                                customer = customer_match.group(1) if customer_match else "Unknown"
                                
                                mismatch_data.append({
                                    'disp_no': disp_no,
                                    'current_route': current_route,
                                    'correct_route': correct_route,
                                    'customer': customer
                                })
                    except Exception as e:
                        print(f"Error parsing mismatch data: {e}")
                        continue

        if not mismatch_data:
            messagebox.showerror("No Mismatches", f"No route mismatches selected to process.\n\nTotal mismatches found: {total_mismatches}\nSelected: {selected_count}")
            return

        # Confirm with user
        confirm_msg = f"Found {selected_count} selected route mismatches to fix (out of {total_mismatches} total).\n\n"
        confirm_msg += "This will automatically:\n"
        confirm_msg += "1. Open Pickup Request for each dispatch\n"
        confirm_msg += "2. Enter the location and dispatch number\n"
        confirm_msg += "3. Click Assign Route and enter the correct route\n"
        confirm_msg += "4. Click OK to save\n\n"
        confirm_msg += "Continue?"
        
        if not messagebox.askyesno("Confirm Stop Movement", confirm_msg):
            return

        # Disable the button during processing
        self.stop_movement_button.configure(state="disabled")
        
        # Disable Start and Stop buttons to prevent accidental activation
        self.sort_start_button.configure(state="disabled")
        self.sort_stop_button.configure(state="disabled")
        
        # Unbind Enter key from location entry to prevent accidental route mismatch activation
        self.sort_location_entry.unbind("<Return>")
        
        # Clear all checkboxes when starting auto move (only if UI exists)
        if hasattr(self, 'select_all_var'):
            self.select_all_var.set(False)
        if hasattr(self, 'result_frames'):
            for result in self.result_frames:
                if result['checkbox'] is not None:
                    result['checkbox_var'].set(False)
        
        # Hide checkboxes from GUI when auto move starts
        self.hide_checkboxes()
        print("ðŸ” Checkboxes hidden for auto move process")
        
        # Don't clear results immediately - wait for first successful move
        # Show status without clearing results yet
        self.sort_status_label.configure(text="Processing Stop Movement")
        
        # Start breathing animation for stop movement
        self.sort_breathing_alpha = 1.0
        self.sort_breathing_direction = -1
        self._breathing_animation()
        
        # Initialize progress bar properly for auto move
        self.update_custom_progress(0)
        self.update_idletasks()  # Force UI update

        # Launch the automation in a separate thread
        threading.Thread(
            target=self._run_stop_movement_worker,
            args=(mismatch_data,),
            daemon=True
        ).start()



    def _run_stop_movement_worker(self, mismatch_data):
        """Worker thread for stop movement automation"""
        try:
            # Initialize Selenium driver in headless mode (same as other functions)
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--headless")  # Headless mode like other functions
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-gpu")
            
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # Use the same login approach as other functions - with MFA callbacks
            login_helper = CarrierSeleniumMonitor(
                username=self.username,
                password=self.password,
                dispatch_location=self.sort_location_entry.get().strip().upper(),
                stop_event=threading.Event(),
                message_types=[],
                mfa_factor_callback=self.prompt_mfa_factor,
                mfa_code_callback=self.prompt_mfa_code
            )
            
            # Navigate and login using the proven method
            wait = WebDriverWait(driver, 30)
            login_helper.sso_navigate(
                driver, wait,
                "https://internal.example.com",
                self.username, self.password
            )
            
            print("âœ… Login successful using existing MFA callback system!")
            
            # Process each mismatch
            total_mismatches = len(mismatch_data)
            for i, mismatch in enumerate(mismatch_data):
                try:
                    print(f"Processing mismatch {i+1}/{total_mismatches}: Disp#{mismatch['disp_no']}", flush=True)
                    
                    # Update status and progress
                    status_text = f"Moving Disp#{mismatch['disp_no']} from Route {mismatch['current_route']} to {mismatch['correct_route']}"
                    self.sort_status_label.configure(text=status_text)
                    
                    # Ensure breathing animation continues for individual stop movements
                    if not hasattr(self, 'sort_breathing_alpha'):
                        self.sort_breathing_alpha = 1.0
                        self.sort_breathing_direction = -1
                        self._breathing_animation()
                    
                    progress = ((i + 1) / total_mismatches) * 100
                    self.update_custom_progress(progress)
                    self.update_idletasks()  # Force UI update for smooth progress bar
                    
                    # Process this mismatch with retry logic (first mismatch opens new form, subsequent ones reuse existing form)
                    is_first_mismatch = (i == 0)
                    success = False
                    max_retries = 3
                    
                    for retry_attempt in range(max_retries):
                        if retry_attempt > 0:
                            print(f"ðŸ”„ Retry attempt {retry_attempt + 1}/{max_retries} for Disp#{mismatch['disp_no']}", flush=True)
                            status_text = f"Retry {retry_attempt + 1}/{max_retries}: Moving Disp#{mismatch['disp_no']} from Route {mismatch['current_route']} to {mismatch['correct_route']}"
                            self.sort_status_label.configure(text=status_text)
                            time.sleep(2)  # Brief delay before retry
                        
                        success = self._process_single_stop_movement(driver, mismatch, is_first_mismatch)
                        
                        if success:
                            break
                        elif retry_attempt < max_retries - 1:  # Don't sleep on the last attempt
                            print(f"âš ï¸ Attempt {retry_attempt + 1} failed, waiting before retry...", flush=True)
                            time.sleep(3)  # Wait 3 seconds before retry
                    
                    # Add status update to results (without checkboxes)
                    if success:
                        retry_text = f" (attempt {retry_attempt + 1})" if retry_attempt > 0 else ""
                        self.add_result(f"âœ… Successfully moved Disp#{mismatch['disp_no']}: Route {mismatch['current_route']} â†’ {mismatch['correct_route']}{retry_text}", show_checkbox=False)
                    else:
                        self.add_result(f"âŒ Failed to move Disp#{mismatch['disp_no']}: Route {mismatch['current_route']} â†’ {mismatch['correct_route']} after {max_retries} attempts", show_checkbox=False)
                    
                    # Small delay between operations
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Error processing Disp#{mismatch['disp_no']}: {e}", flush=True)
                    continue
            
            # Final update
            self.update_custom_progress(100.0)
            self.update_idletasks()  # Force UI update for final progress
            self.sort_status_label.configure(text="Stop Movement Complete", text_color=BRAND_ORANGE)
            
            # Stop breathing animation by resetting alpha
            self.sort_breathing_alpha = 1.0
            
            # Show completion message (without checkboxes)
            self.add_result("", show_checkbox=False)
            self.add_result("=" * 60, show_checkbox=False)
            self.add_result("âœ… STOP MOVEMENT COMPLETE", show_checkbox=False)
            self.add_result("=" * 60, show_checkbox=False)
            self.add_result(f"Processed {total_mismatches} route mismatches.", show_checkbox=False)
            
        except Exception as e:
            print(f"Error in stop movement automation: {e}", flush=True)
            self.sort_status_label.configure(text="Error", text_color="red")
            self.add_result(f"âŒ Error: {e}", show_checkbox=False)
        finally:
            # Re-enable the Auto Move button
            if hasattr(self, 'auto_move_button'):
                self.auto_move_button.configure(state="normal")
            
            # Re-enable Start and Stop buttons
            if hasattr(self, 'sort_start_button'):
                self.sort_start_button.configure(state="normal")
            if hasattr(self, 'sort_stop_button'):
                self.sort_stop_button.configure(state="disabled")  # Keep stop disabled until next process starts
            
            # Re-bind Enter key to location entry
            if hasattr(self, 'sort_location_entry'):
                self.sort_location_entry.bind(
                    "<Return>",
                    lambda e: self.start_sort()
                )
            
            # Cleanup driver
            try:
                if 'driver' in locals():
                    driver.quit()
            except:
                pass

    def _login_to_cxpc(self, driver):
        """Login to CXPC application using the same pattern as other functions"""
        try:
            # Create login helper using the same pattern as other functions
            login_helper = CarrierSeleniumMonitor(
                username=self.username,
                password=self.password,
                dispatch_location="",
                stop_event=threading.Event(),
                message_types=[],
                mfa_factor_callback=self.prompt_mfa_factor,
                mfa_code_callback=self.prompt_mfa_code
            )
            
            # Use the same login URL as other functions
            login_url = "https://internal.example.com"
            
            # Use the same sso_navigate method as other functions
            wait = WebDriverWait(driver, 10)
            login_helper.sso_navigate(driver, wait, login_url, self.username, self.password)
            
            print("âœ… Successfully logged in using same credentials as other functions", flush=True)
            
        except Exception as e:
            print(f"âŒ Login failed: {e}", flush=True)
            raise

    def _process_single_stop_movement(self, driver, mismatch, is_first_mismatch=True):
        """Process a single stop movement for one dispatch - PROVEN WORKING IMPLEMENTATION"""
        try:
            print(f"ðŸ” Starting process for Disp#{mismatch['disp_no']} (first_mismatch={is_first_mismatch})", flush=True)
            
            # Wait for page to fully load
            time.sleep(1)  # Reduced from 2 seconds
            
            if is_first_mismatch:
                # Only click "Pickup Request" for the first mismatch - setup form
                print(f"ðŸ” First mismatch - opening Pickup Request form...", flush=True)
                
                pickup_request = WebDriverWait(driver, 10).until(  # Reduced from 15 seconds
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Pickup Request')]"))
                )
                pickup_request.click()
                time.sleep(1.5)  # Reduced from 3 seconds
                print(f"âœ… Opened Pickup Request form", flush=True)
                
                # Enter location from the UI input field
                location_field = driver.find_element(By.XPATH, "//input[@formcontrolname='dispatchLocation']")
                location_field.clear()
                location_field.send_keys(self.sort_location_entry.get().strip().upper())
                print(f"âœ… Entered location: {self.sort_location_entry.get().strip().upper()}", flush=True)
                
                # Enter dispatch number
                dispatch_field = driver.find_element(By.XPATH, "//input[@formcontrolname='pickupNbr']")
                dispatch_field.clear()
                dispatch_field.send_keys(mismatch['disp_no'])
                dispatch_field.send_keys(Keys.RETURN)
                print(f"âœ… Entered dispatch: {mismatch['disp_no']}", flush=True)
                time.sleep(1)  # Reduced from 2 seconds
            else:
                # For subsequent mismatches, just update the dispatch number
                print(f"ðŸ” Subsequent mismatch - updating dispatch number for {mismatch['disp_no']}", flush=True)
                dispatch_field = driver.find_element(By.XPATH, "//input[@formcontrolname='pickupNbr']")
                dispatch_field.clear()
                dispatch_field.send_keys(mismatch['disp_no'])
                dispatch_field.send_keys(Keys.RETURN)
                print(f"âœ… Updated dispatch: {mismatch['disp_no']}", flush=True)
                time.sleep(1)  # Reduced from 2 seconds
            
            # Click Assign Route to open modal (PROVEN PATTERN)
            print("ðŸ” Clicking 'Assign Route' button...")
            assign_route_button = WebDriverWait(driver, 8).until(  # Reduced from 10 seconds
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Assign Route')]"))
            )
            assign_route_button.click()
            time.sleep(1)  # Reduced from 2 seconds
            
            # Detect Modal with Retry Logic (OPTIMIZED)
            print("ðŸ” Detecting modal...")
            modal_container = None
            max_retries = 5  # Reduced from 10
            retry_count = 0
        
            while retry_count < max_retries and modal_container is None:
                try:
                    modal_container = WebDriverWait(driver, 1).until(  # Reduced from 2 seconds
                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-dialog')]"))
                    )
                    print("âœ… Modal detected!")
                    break
                except:
                    retry_count += 1
                    print(f"ðŸ” Modal not found, retry {retry_count}/{max_retries}...")
                    time.sleep(0.3)  # Reduced from 1 second
                    
                    # Try clicking Assign Route again if modal doesn't appear
                    if retry_count <= 3:  # Reduced from 5 attempts
                        try:
                            print("ðŸ” Retrying Assign Route button click...")
                            assign_route_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Assign Route')]")
                            assign_route_button.click()
                            time.sleep(0.5)  # Reduced from 2 seconds
                        except Exception as retry_error:
                            print(f"ðŸ” Retry click failed: {retry_error}")
            
            if modal_container is None:
                print("âŒ Modal still not detected after 5 retries - this may cause issues")
                return False
            
            # Apply Brute Force Route Setting within Modal Context (PROVEN PATTERN)
            print(f"ðŸ” Setting route to {mismatch['correct_route']} using brute force...")
            
            result = driver.execute_script(f"""
                var modal = arguments[0];
                var input = modal.querySelector('input[class*="p-autocomplete-input"]');
                if (input) {{
                    // Brute force approach (PROVEN TO WORK)
                    var originalEvents = {{}};
                    ['input', 'change', 'keyup', 'keydown', 'focus', 'blur'].forEach(function(eventType) {{
                        originalEvents[eventType] = input['on' + eventType];
                        input['on' + eventType] = null;
                    }});
                    
                    input.disabled = false;
                    input.removeAttribute('disabled');
                    input.readOnly = false;
                    input.removeAttribute('readonly');
                    input.tabIndex = 0;
                    input.style.pointerEvents = 'auto';
                    input.style.display = 'block';
                    input.style.visibility = 'visible';
                    
                    input.value = '{mismatch['correct_route']}';
                    input.setAttribute('value', '{mismatch['correct_route']}');
                    input.defaultValue = '{mismatch['correct_route']}';
                    
                    Object.defineProperty(input, 'value', {{
                        get: function() {{ return '{mismatch['correct_route']}'; }},
                        set: function(val) {{ this.setAttribute('value', '{mismatch['correct_route']}'); }},
                        configurable: true
                    }});
                    
                    ['input', 'change', 'keyup', 'keydown', 'focus', 'blur'].forEach(function(eventType) {{
                        input['on' + eventType] = originalEvents[eventType];
                        var event = new Event(eventType, {{ bubbles: true, cancelable: true }});
                        input.dispatchEvent(event);
                    }});
                    
                    var enterEvent = new KeyboardEvent('keydown', {{
                        key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
                    }});
                    input.dispatchEvent(enterEvent);
                    
                    return input.value;
                }}
                return null;
            """, modal_container)
            
            if result != mismatch['correct_route']:
                print(f"âŒ FAILED: Expected '{mismatch['correct_route']}', got '{result}'")
                return False
            
            print(f"âœ… SUCCESS! Route set to '{result}'")
            
            # Handle autocomplete dropdown (OPTIMIZED)
            print("ðŸ” Checking for autocomplete dropdown...")
            try:
                time.sleep(0.3)  # Reduced from 1 second
                autocomplete_items = driver.find_elements(By.XPATH, f"//li[contains(@class, 'p-autocomplete-item') and contains(text(), '{mismatch['correct_route']}')]")
                if autocomplete_items:
                    print(f"ðŸ” Found {len(autocomplete_items)} autocomplete items, clicking the first one...")
                    for item in autocomplete_items:
                        if item.is_displayed():
                            try:
                                item.click()
                                print(f"âœ… Clicked autocomplete item: '{item.text}'")
                                time.sleep(0.3)  # Reduced from 1 second
                                break
                            except Exception as e:
                                try:
                                    driver.execute_script("arguments[0].click();", item)
                                    print(f"âœ… Clicked autocomplete item via JavaScript: '{item.text}'")
                                    time.sleep(0.3)  # Reduced from 1 second
                                    break
                                except:
                                    continue
                else:
                    print("ðŸ” No autocomplete items found, proceeding to OK button")
            except Exception as auto_error:
                print(f"ðŸ” Autocomplete handling failed: {auto_error}")
            
            # Click OK Button (OPTIMIZED)
            print("ðŸ” Clicking OK button...")
            try:
                time.sleep(0.3)  # Reduced from 1 second
                
                # Find OK button within modal
                ok_button = modal_container.find_element(By.XPATH, ".//button[contains(@class, 'pickupButtons')]")
                print("ðŸ” Found OK button in modal")
                
                # Try clicking OK button with fallbacks
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", ok_button)
                    time.sleep(0.2)  # Reduced from 0.5 seconds
                    ok_button.click()
                    print("âœ… OK button clicked with standard approach!")
                except Exception as e:
                    try:
                        driver.execute_script("arguments[0].click();", ok_button)
                        print("âœ… OK button clicked via JavaScript!")
                    except Exception as js_e:
                        print(f"âŒ JavaScript click also failed: {js_e}")
                        return False
            except Exception as ok_error:
                print(f"âŒ OK button process failed: {ok_error}")
                return False
            
            # Handle specific popups that may appear after route assignment
            print("ðŸ” Checking for specific popups (Route blocked, Route not signed on)...")
            time.sleep(0.5)
            
            # Check for "Route has been Blocked" popup
            blocked_route_patterns = [
                "//div[contains(text(), 'Route') and contains(text(), 'has been Blocked')]",
                "//div[contains(text(), 'Route') and contains(text(), 'Blocked')]",
                "//span[contains(text(), 'Route') and contains(text(), 'has been Blocked')]",
                "//span[contains(text(), 'Route') and contains(text(), 'Blocked')]",
                "//p[contains(text(), 'Route') and contains(text(), 'has been Blocked')]",
                "//p[contains(text(), 'Route') and contains(text(), 'Blocked')]"
            ]
            
            blocked_popup_found = False
            for pattern in blocked_route_patterns:
                try:
                    blocked_elements = driver.find_elements(By.XPATH, pattern)
                    for element in blocked_elements:
                        if element.is_displayed():
                            print(f"ðŸ” Found 'Route has been Blocked' popup: {element.text}")
                            blocked_popup_found = True
                            break
                    if blocked_popup_found:
                        break
                except:
                    continue
            
            # Check for "Route not signed on" popup
            route_not_signed_patterns = [
                "//div[contains(text(), 'You are assigning PU Stops to a Route that is not signed on')]",
                "//div[contains(text(), 'Route that is not signed on')]",
                "//span[contains(text(), 'You are assigning PU Stops to a Route that is not signed on')]",
                "//span[contains(text(), 'Route that is not signed on')]",
                "//p[contains(text(), 'You are assigning PU Stops to a Route that is not signed on')]",
                "//p[contains(text(), 'Route that is not signed on')]"
            ]
            
            route_not_signed_found = False
            for pattern in route_not_signed_patterns:
                try:
                    route_not_signed_elements = driver.find_elements(By.XPATH, pattern)
                    for element in route_not_signed_elements:
                        if element.is_displayed():
                            print(f"ðŸ” Found 'Route not signed on' popup: {element.text}")
                            route_not_signed_found = True
                            break
                    if route_not_signed_found:
                        break
                except:
                    continue
            
            # Handle "Route not signed on" popup with Continue button
            if route_not_signed_found:
                continue_patterns = [
                    "//button[contains(text(), 'Continue')]",
                    "//button[contains(@class, 'pickupButtons') and contains(text(), 'Continue')]",
                    "//span[text()='Continue']/ancestor::button[1]",
                    "//button[contains(@class, 'p-button') and contains(text(), 'Continue')]"
                ]
                
                for pattern in continue_patterns:
                    try:
                        continue_buttons = driver.find_elements(By.XPATH, pattern)
                        for btn in continue_buttons:
                            if btn.is_displayed() and btn.is_enabled():
                                try:
                                    driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                                    time.sleep(0.2)
                                    btn.click()
                                    print(f"âœ… Handled 'Route not signed on' popup: Clicked 'Continue' button")
                                    time.sleep(0.5)
                                    break
                                except Exception as click_error:
                                    try:
                                        driver.execute_script("arguments[0].click();", btn)
                                        print(f"âœ… Handled 'Route not signed on' popup: Clicked 'Continue' button via JavaScript")
                                        time.sleep(0.5)
                                        break
                                    except Exception as js_error:
                                        print(f"ðŸ” Continue button JavaScript click failed: {js_error}")
                                        continue
                        if btn.is_displayed() and btn.is_enabled():
                            break
                    except Exception as pattern_error:
                        print(f"ðŸ” Continue pattern {pattern} failed: {pattern_error}")
                        continue
            
            # Handle "Route has been Blocked" popup with Yes/OK button
            if blocked_popup_found:
                yes_ok_patterns = [
                    "//button[contains(text(), 'Yes')]",
                    "//button[contains(text(), 'OK')]", 
                    "//button[contains(@class, 'p-confirm-dialog-accept')]",
                    "//button[contains(@class, 'p-button-success')]",
                    "//span[text()='Yes']/ancestor::button[1]",
                    "//span[text()='OK']/ancestor::button[1]",
                    "//button[contains(@class, 'p-button') and contains(text(), 'Yes')]",
                    "//button[contains(@class, 'p-button') and contains(text(), 'OK')]"
                ]
                
                for pattern in yes_ok_patterns:
                    try:
                        yes_ok_buttons = driver.find_elements(By.XPATH, pattern)
                        for btn in yes_ok_buttons:
                            if btn.is_displayed() and btn.is_enabled():
                                try:
                                    driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                                    time.sleep(0.2)
                                    btn.click()
                                    print(f"âœ… Handled 'Route has been Blocked' popup: Clicked '{btn.text}' button")
                                    time.sleep(0.5)
                                    break
                                except Exception as click_error:
                                    try:
                                        driver.execute_script("arguments[0].click();", btn)
                                        print(f"âœ… Handled 'Route has been Blocked' popup: Clicked '{btn.text}' button via JavaScript")
                                        time.sleep(0.5)
                                        break
                                    except Exception as js_error:
                                        print(f"ðŸ” Yes/OK button JavaScript click failed: {js_error}")
                                        continue
                        if btn.is_displayed() and btn.is_enabled():
                            break
                    except Exception as pattern_error:
                        print(f"ðŸ” Yes/OK pattern {pattern} failed: {pattern_error}")
                        continue
            
            # Handle SUCCESS popup and Click "Close" button (OPTIMIZED)
            print("ðŸ” Looking for success popup and Close button...")
            time.sleep(0.5)  # Reduced from 2 seconds
            
            try:
                # Look for Close button with the specific span text
                close_patterns = [
                    "//span[@class='p-button-label' and text()='Close']",
                    "//button[.//span[text()='Close']]",
                    "//span[text()='Close']",
                    "//button[contains(@class, 'p-button') and .//span[contains(text(), 'Close')]]"
                ]
                
                close_clicked = False
                for pattern in close_patterns:
                    try:
                        close_elements = driver.find_elements(By.XPATH, pattern)
                        for close_element in close_elements:
                            if close_element.is_displayed():
                                print(f"ðŸ” Found Close element using pattern: {pattern}")
                                try:
                                    # If it's a span, find the parent button
                                    if close_element.tag_name == 'span':
                                        close_button = close_element.find_element(By.XPATH, "./ancestor::button[1]")
                                    else:
                                        close_button = close_element
                                    
                                    close_button.click()
                                    print(f"âœ… Clicked Close button successfully!")
                                    close_clicked = True
                                    time.sleep(0.5)  # Reduced from 2 seconds
                                    break
                                except Exception as e:
                                    print(f"ðŸ” Standard close click failed: {e}")
                                    try:
                                        if close_element.tag_name == 'span':
                                            close_button = close_element.find_element(By.XPATH, "./ancestor::button[1]")
                                        else:
                                            close_button = close_element
                                        driver.execute_script("arguments[0].click();", close_button)
                                        print(f"âœ… Clicked Close button via JavaScript!")
                                        close_clicked = True
                                        time.sleep(0.5)  # Reduced from 2 seconds
                                        break
                                    except Exception as js_e:
                                        print(f"âŒ JavaScript close click also failed: {js_e}")
                                        continue
                        if close_clicked:
                            break
                    except Exception as pattern_error:
                        print(f"ðŸ” Pattern {pattern} failed: {pattern_error}")
                        continue
                
                if not close_clicked:
                    print("âš ï¸  Could not find or click Close button - but assignment may have succeeded")
                
            except Exception as close_error:
                print(f"ðŸ” Close button handling failed: {close_error}")
            
            print(f"âœ… Successfully processed Disp#{mismatch['disp_no']}: Route {mismatch['current_route']} â†’ {mismatch['correct_route']}")
            return True
            
        except Exception as e:
            print(f"âŒ Error processing Disp#{mismatch['disp_no']}: {e}")
            
            # Take a screenshot for debugging
            try:
                screenshot_path = f"error_disp_{mismatch['disp_no']}_{int(time.time())}.png"
                driver.save_screenshot(screenshot_path)
                print(f"ðŸ“¸ Screenshot saved: {screenshot_path}")
            except Exception as screenshot_error:
                print(f"âŒ Failed to take screenshot: {screenshot_error}")
            
            return False
    
    def _convert_to_uppercase_sort(self, event=None):
        """
        Keep the sort-location entry to max 4 uppercase chars.
        """
        txt = self.sort_location_entry.get().upper()[:4]
        self.sort_location_entry.delete(0, ctk.END)
        self.sort_location_entry.insert(0, txt)

    def _create_auto_sort_ui(self, sort_tab):
        """Create the Auto-Sort UI components"""
        from customtkinter import CTkProgressBar
        import sys
        import os

        sort_frame = ctk.CTkFrame(
            sort_tab,
            fg_color="black",
            border_width=2,
            border_color=BRAND_ORANGE
        )
        sort_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        # Top box: Controls and progress
        top_box = ctk.CTkFrame(
            sort_frame,
            fg_color="black",
            border_width=2,
            border_color=BRAND_ORANGE
        )
        top_box.pack(padx=10, pady=(10,5), fill="x")

        # Centered header
        ctk.CTkLabel(
            top_box,
            text="Sort Regular Pickups",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(pady=(10,15))

        # Controls row: Location entry + Start/Stop buttons
        ctrl = ctk.CTkFrame(top_box, fg_color="black")
        ctrl.pack(padx=10, pady=(0,10))

        ctk.CTkLabel(
            ctrl,
            text="Location:",
            text_color=BRAND_ORANGE,
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left", padx=(0,5))

        self.sort_location_entry = ctk.CTkEntry(
            ctrl,
            width=100,
            fg_color="white",     # white background
            text_color="black"    # black text
        )
        self.sort_location_entry.pack(side="left")
        # â—€ Bind to uppercase + max-4 helper
        self.sort_location_entry.bind(
            "<KeyRelease>",
            lambda e: self._convert_to_uppercase_sort()
        )
        # â—€ Bind Enter key to start sort function
        self.sort_location_entry.bind(
            "<Return>",
            lambda e: self.start_sort()
        )

        self.sort_start_button = ctk.CTkButton(
            ctrl,
            text="Start",
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self.start_sort
        )
        self.sort_start_button.pack(side="left", padx=(5,5))

        self.sort_stop_button = ctk.CTkButton(
            ctrl,
            text="Stop",
            fg_color="red",
            hover_color="darkred",
            command=self.stop_sort,
            state="disabled"  # Initially disabled until process starts
        )
        self.sort_stop_button.pack(side="left", padx=(0,5))

        # Custom progress bar with Carrier truck and road
        self.progress_frame = ctk.CTkFrame(top_box, fg_color="black", height=60)
        self.progress_frame.pack(padx=10, pady=(10,5), fill="x")
        self.progress_frame.pack_propagate(False)
        
        # Road background (black with orange lines)
        self.road_canvas = tk.Canvas(
            self.progress_frame,
            bg="black",  # Black road to match UI
            height=40,
            highlightthickness=0
        )
        self.road_canvas.pack(pady=10, fill="x", padx=5)
        
        # Create progress overlay canvas
        self.progress_canvas = tk.Canvas(
            self.progress_frame,
            bg="black",
            height=40,
            highlightthickness=0
        )
        self.progress_canvas.place(x=5, y=10, relwidth=0, height=40)
        
        # Draw road lines (initially empty)
        self.draw_road_lines(0)
        
        # Load Carrier truck logo
        try:
            from PIL import Image, ImageTk
            # Use bundled logo file (works under PyInstaller)
            base_dir = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
            truck_logo_path = os.path.join(base_dir, "TRUCK_LOGO.png")
            truck_img = Image.open(truck_logo_path)
            # Resize to fit progress bar height
            truck_img = truck_img.resize((50, 30), Image.LANCZOS)
            self.truck_photo = ImageTk.PhotoImage(truck_img)
            
            # Create truck on main road canvas (initially hidden)
            self.truck_id = self.road_canvas.create_image(
                10, 20,  # Start position
                image=self.truck_photo,
                anchor="w",
                state="hidden"  # Initially hidden
            )
        except Exception as e:
            print(f"Could not load truck logo: {e}")
            # Fallback: create a simple truck shape
            self.truck_id = self.road_canvas.create_rectangle(
                10, 10, 60, 30,
                fill=BRAND_ORANGE,
                outline="white",
                state="hidden"  # Initially hidden
            )
        
        # Truck breathing animation variables
        self.truck_breathing_active = False
        self.truck_breathing_alpha = 1.0
        self.truck_breathing_direction = -1
        
        # Status label
        self.sort_status_label = ctk.CTkLabel(
            top_box,
            text="Ready",
            font=ctk.CTkFont(size=12),
            text_color="white"
        )
        self.sort_status_label.pack(pady=(5,10))
        
        # Progress text overlay
        self.progress_text = ctk.CTkLabel(
            self.progress_frame,
            text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=BRAND_ORANGE
        )
        self.progress_text.place(relx=0.5, rely=0.5, anchor="center")

        # Results area (scrollable)
        results_frame = ctk.CTkFrame(
            sort_frame,
            fg_color="black",
            border_width=2,
            border_color=BRAND_ORANGE
        )
        results_frame.pack(padx=10, pady=(5,10), fill="both", expand=True)

        # Results header with copy button and auto move button
        results_header = ctk.CTkFrame(results_frame, fg_color="black")
        results_header.pack(fill="x", padx=10, pady=(10,5))

        ctk.CTkLabel(
            results_header,
            text="Results:",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=BRAND_ORANGE
        ).pack(side="left")

        # Auto Move button (initially hidden until Shift+F7 is pressed)
        self.auto_move_button = ctk.CTkButton(
            results_header,
            text="Auto Move",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self.start_auto_move
        )
        # Don't pack it initially - it will be shown when Shift+F7 is pressed

        # Copy button
        self.copy_results_button = ctk.CTkButton(
            results_header,
            text="Copy Results",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            command=self.copy_results_to_clipboard
        )
        self.copy_results_button.pack(side="right")

        # Select All checkbox at the top of results
        select_all_frame = ctk.CTkFrame(results_frame, fg_color="black")
        select_all_frame.pack(fill="x", padx=10, pady=(5,0))

        self.select_all_var = tk.BooleanVar()
        self.select_all_checkbox = ctk.CTkCheckBox(
            select_all_frame,
            text="Select All",
            variable=self.select_all_var,
            command=self.toggle_select_all,
            fg_color=BRAND_ORANGE,
            hover_color=BRAND_PURPLE,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.select_all_checkbox.pack(side="left", padx=(0, 10))

        # Scrollable results area
        self.results_scrollable = ctk.CTkScrollableFrame(
            results_frame,
            fg_color="black",
            width=540,
            height=300
        )
        self.results_scrollable.pack(padx=10, pady=(5,10), fill="both", expand=True)

        # Initialize results storage
        self.result_frames = []
        self.sort_stop_event = threading.Event()

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def on_closing(self):
        # ask for confirmation first
        if not messagebox.askyesno("Confirm", "Are you sure you want to close out the application?"):
            return

        # Disable logging to prevent spam during shutdown
        import logging
        logging.getLogger().setLevel(logging.CRITICAL)
        
        # Stop all scheduled Tkinter "after" tasks
        try:
            for job_id in list(self.tk.call('after', 'info')):
                self.after_cancel(job_id)
        except:
            pass

        # Signal all monitors/threads to stop, but don't wait on or block for browser shutdowns here
        for info in list(self.running_monitors.values()):
            try:
                info["stop_event"].set()
            except:
                pass
                
        for info in list(self.running_clear_monitors.values()):
            try:
                info["stop_event"].set()
            except:
                pass
        
        # Signal all audit monitors to stop as well so they don't keep working during shutdown
        for info in list(getattr(self, "running_audit_monitors", {}).values()):
            try:
                info.get("stop_event").set()
            except:
                pass
        
        # Signal SoF+ monitor
        if hasattr(self, "sof_monitor_stop_event"):
            try:
                self.sof_monitor_stop_event.set()
            except:
                pass
        
        # Best-effort: browsers and chromedrivers will be terminated below via cleanup_processes/sys.exit

        # Stop RoutR Chrome Extension server so extension stops working when ROUTR is closed
        self._stop_routr_extension_server()

        # Clean up widget window if it exists
        if hasattr(self, 'widget_window') and self.widget_window:
            try:
                self.widget_panel_visible = False
                self.widget_window.destroy()
                self.widget_window = None
            except:
                pass

        # Clean up widget button canvas if it exists
        if hasattr(self, 'widget_button_canvas') and self.widget_button_canvas:
            try:
                self.widget_button_canvas.destroy()
            except:
                pass

        # Skip explicit child-process cleanup here to keep exit instant
        # (chromedriver/Chrome will be left for the OS to clean up)
        pass
        
        # Clean up console handler if it exists
        if hasattr(self, 'console_handler'):
            try:
                root_logger = logging.getLogger()
                root_logger.removeHandler(self.console_handler)
                self.console_handler.close()
            except:
                pass
        
        # Restore original stdout if console was used
        if hasattr(self, 'original_stdout'):
            try:
                import sys
                if hasattr(self, 'console_stdout'):
                    self.console_stdout.close()
                sys.stdout = self.original_stdout
            except:
                pass

        # Properly quit and destroy to prevent ghost processes
        try:
            self.quit()
        except:
            pass
        try:
            self.destroy()
        except:
            pass
        
        # Force exit if still running (prevents ghost processes)
        import sys
        try:
            sys.exit(0)
        except:
            pass


if __name__ == "__main__":
    app = CarrierMonitorApp()
    # Protocol is already set in __init__, no need to set it again
    app.mainloop()

#--------------
#venv\Scripts\activate.bat
#pyinstaller --clean ROUTR.spec





"""TMS appointment uploader: writes booked appointments from the shared master back into the TMS.

Portfolio example. Screen IDs are simplified; addresses and paths are placeholders.
Status: pre-production. The whole form is filled; Submit is pressed by a person, who confirms
the result at the prompt, until the automated Submit is signed off.

Flow
  1. Read master rows that a booking bot finished (APPT Scheduling C/F, Process Stage 2)
  2. Initial / Reschedule: open the Load in the TMS and fill the appointment form
  3. Every type: append a result row and mail the owning load-planning team
  4. Mark the row done: upload status C, Process Stage 3
"""
import re
from datetime import date, time
from pathlib import Path

import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

UPLOAD_TYPES = ("initial", "reschedule")
READY = ("C", "F")
REQUEST_TYPE_MAP = {"I": "Initial", "R": "Reschedule", "M": "Modification", "C": "Cancel"}
UNLOAD_MAP = {"Live": "N", "Drop": "Y"}
RESCHEDULE_ROUNDS = ["2nd", "3rd", "4th", "5th"]   # 1st is Initial-only

INITIAL_REQUIRED = ["1st Appt.Confirmation No.", "1st Appt.Delivery D/T", "1st Appt.Delivery T/M",
                    "1st Appt.Request D/T", "1st Appt.Request T/M"]


# ============================================================
# Master: what is ready to upload
# ============================================================

def latest_reschedule_round(get) -> str | None:
    """The highest of 2nd..5th that has a requested date."""
    latest = None
    for rnd in RESCHEDULE_ROUNDS:
        if get(f"{rnd} Requested Appt.D/T"):
            latest = rnd
    return latest


def get_pending(master_path: Path) -> list[dict]:
    """Outbound rows a booking bot finished. Initial/Reschedule feed the TMS form, so a row
    missing any form field is skipped; Modification/Cancel never touch the form."""
    ws = openpyxl.load_workbook(master_path).active
    col = {str(c.value): c.column - 1 for c in ws[1] if c.value}
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        get = lambda name: row[col[name]] if name in col else None
        if get("DEPT") != "Outbound" or not get("Load ID"):
            continue
        if get("APPT Scheduling") not in READY or str(get("Process Stage")) != "2":
            continue
        rtype = REQUEST_TYPE_MAP.get(get("Request Type"), get("Request Type"))
        appt = {"type": rtype.lower(), "load_id": get("Load ID"), "plan_id": get("Plan ID"),
                "unload_type": UNLOAD_MAP.get(get("UNLOAD TYPE"), "")}
        if rtype == "Initial":
            if not all(get(c) for c in INITIAL_REQUIRED):
                continue
            appt.update(appt_no=get("1st Appt.Confirmation No."), appt_date=get("1st Appt.Delivery D/T"),
                        appt_time=get("1st Appt.Delivery T/M"), request_date=get("1st Appt.Request D/T"),
                        request_time=get("1st Appt.Request T/M"))
        elif rtype == "Reschedule":
            rnd = latest_reschedule_round(get)
            needed = [f"{rnd} Requested Appt.T/M", f"{rnd} Appt.Request D/T",
                      f"{rnd} Appt.Request T/M", f"{rnd} Reschedule Reason Code"] if rnd else []
            if rnd is None or not all(get(c) for c in needed):
                continue
            appt.update(appt_date=get(f"{rnd} Requested Appt.D/T"), appt_time=get(f"{rnd} Requested Appt.T/M"),
                        request_date=get(f"{rnd} Appt.Request D/T"), request_time=get(f"{rnd} Appt.Request T/M"),
                        reason_code=get(f"{rnd} Reschedule Reason Code"))
        out.append(appt)
    return out


def mark_uploaded(master_path: Path, load_id, request_type: str) -> None:
    wb = openpyxl.load_workbook(master_path)
    ws = wb.active
    col = {str(c.value): c.column - 1 for c in ws[1] if c.value}
    for row in ws.iter_rows(min_row=2):
        if str(row[col["Load ID"]].value) == str(load_id):
            row[col["Upload Status"]].value = "C"
            row[col["Process Stage"]].value = "3"
            if request_type == "reschedule":
                row[col["Reschedule Y/N"]].value = None   # this reschedule is closed out
            break
    wb.save(master_path)


# ============================================================
# Recipients: Plan ID → location → load-planning team
# ============================================================

LOCATION_EMAIL = {"Location A": "lp-a@example.com", "Location B": "lp-b@example.com"}


def plan_id_to_location(rows) -> dict:
    """The region sheet lists a location once, then its plan IDs on the rows below
    (merged cells): carry the last seen location down. 'Name (CODE)' → 'Name'."""
    table, current = {}, None
    for location, plan_id in rows:
        if location:
            current = re.sub(r"\s*\(.*?\)\s*", "", location).strip()
        if plan_id is not None and current is not None:
            table[str(plan_id)] = current
    return table


# ============================================================
# TMS screens
# ============================================================

class TmsUploader:
    def __init__(self, driver):
        self.driver = driver

    def switch_to_result(self):
        # The TMS nests the content: shell frame → result frame
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "shell_content"))
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "resultFrame"))

    def _set_date(self, field_id, value: date):
        el = self.driver.find_element(By.ID, field_id)
        el.clear()
        el.send_keys(value.strftime("%#m/%#d/%Y"))

    def _set_time(self, field_id, value: time):
        """The time boxes are custom widgets that keep the value from the Initial registration.
        send_keys would append to it, and a blur mid-typing makes the page's checkTimeFormat
        alert 'Invalid Time'. Replace the value at once and fire a single change event."""
        el = self.driver.find_element(By.ID, field_id)
        self.driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
            el, value.strftime("%H:%M"))

    def fill_appointment(self, appt: dict) -> None:
        self.switch_to_result()
        Select(self.driver.find_element(By.ID, "dropYN")).select_by_value(appt["unload_type"])
        self.driver.find_element(By.NAME, "apptNumber").send_keys(str(appt["appt_no"]))
        for date_id, time_id, d, t in (("fromDate", "fromHour", appt["appt_date"], appt["appt_time"]),
                                       ("toDate", "toHour", appt["appt_date"], appt["appt_time"]),
                                       ("reqDate", "reqHour", appt["request_date"], appt["request_time"]),
                                       ("sugDate", "sugHour", appt["appt_date"], appt["appt_time"])):
            self._set_date(date_id, d)
            self._set_time(time_id, t)
        confirm_submit(appt["load_id"])


def confirm_submit(load_id) -> None:
    """Until the automated Submit is signed off, a person presses it and confirms here.
    Anything but Enter raises, so the row is never marked uploaded."""
    if input("Submitted? Enter = success, n = failed: ").strip().lower() == "n":
        raise RuntimeError(f"TMS submit reported as failed for load {load_id}")


if __name__ == "__main__":
    sheet = [("Location A (C01)", 101), (None, 102), ("Location B", 201)]
    print(plan_id_to_location(sheet))   # {'101': 'Location A', '102': 'Location A', '201': 'Location B'}

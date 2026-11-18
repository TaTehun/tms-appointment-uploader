# TMS Appointment Uploader

A Python bot that writes booked delivery appointments from a shared Excel master back into the TMS,
then reports each result to the owning load-planning team.

> **Visual overview:** [project-overview.html](project-overview.html)

**Status: pre-production.** Not yet called by the Request Tracker. The appointment form is filled automatically;
a person presses Submit and confirms at the prompt until the automated Submit is signed off.

---

## The team system

This repo is one part of a delivery-appointment automation system built by a small team.
Every program connects only through one shared Excel master file: no program calls another directly.

```
① Request intake      4 request mailboxes → shared master            (mine)
② Date distribution   Available RDD, max 4 bookings/day per customer+DC (mine)
③ Booking             one bot per customer portal, or by email        (mine: Retailer A, Retailer B, email bot · team: 5 portals)
⑤ Confirmation        confirmation mails → master                     (team)
⑦ Upload              confirmed appointments → TMS / WMS              (mine: TMS uploader · team: WMS)
```

| Part | Owner | Repo |
|------|-------|------|
| Request Tracker (parent) + Intake + Distribution | Taehun Jang | [appointment-request-pipeline](https://github.com/TaTehun/appointment-request-pipeline) |
| Portal booking bot, Retailer A | Taehun Jang | [delivery-appointment-scheduler](https://github.com/TaTehun/delivery-appointment-scheduler) |
| Email booking bot | Taehun Jang | [email-appointment-request](https://github.com/TaTehun/email-appointment-request) |
| Portal booking bot, Retailer B (pre-production) | Taehun Jang | [dock-appointment-booking](https://github.com/TaTehun/dock-appointment-booking) |
| TMS uploader (pre-production) | Taehun Jang | [tms-appointment-uploader](https://github.com/TaTehun/tms-appointment-uploader) |
| 5 other portal bots, overseas/truck email bots, confirmation polling, WMS upload | Teammates | private |

---

## What it does

1. **Reads ready rows**: Outbound rows where a booking bot finished (APPT Scheduling `C`/`F`, Process Stage `2`).
   Initial/Reschedule rows missing any form field are skipped until the field is filled.
2. **Opens the Load in the TMS**: Initial goes Plan → Load → Stops; Reschedule goes straight to the Load list. Both open the
   drop stop's appointment form.
3. **Fills the form**: unload type, confirmation number, and four date/time pairs. Time fields are set through JS to avoid the
   page's on-blur validation.
4. **Reports**: appends a row to the result file for every request type and mails the team that owns the Load's Plan ID.
5. **Closes the row**: upload status `C`, Process Stage `3`; a reschedule also clears Reschedule Y/N.

---

## Project Structure (private repo)

```
tms-appointment-uploader/
├── main.py                  # entry point: read, upload, notify, close out
├── processors/
│   ├── tms_upload.py        # TMS screens (Selenium, nested frames)
│   ├── outbound_master.py   # ready rows / close-out
│   ├── result_mapper.py     # result file columns per request type
│   └── region_lookup.py     # Plan ID → location → team address
├── email_pipeline/          # SMTP
└── utils/                   # paths, logging, dates, mail body
```

This public repo holds `main_example.py` and `project-overview.html`.

```
python main_example.py   # Plan ID → location demo
```

---

## Changelog

- **v0.9**: Reads the shared master; result file + team mail for every request type; no browser when only M/C are pending.
- **v0.5**: Reschedule flow; upload status write-back; debug logging.
- **v0.1**: TMS login, frame switching, Plan → Load → Stops → appointment form.

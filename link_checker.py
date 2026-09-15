#!/usr/bin/env python3
"""
Daily Link Checker
-------------------
Reads every hyperlink found in an Excel workbook, checks whether each one
is reachable, and emails a summary report.

Configuration is via environment variables (see README.md):
    EXCEL_FILE_PATH     - path to the .xlsx file (default: links.xlsx)
    SMTP_SENDER_EMAIL   - the Outlook/Office365 address to send FROM
    SMTP_APP_PASSWORD   - app password for that address
    REPORT_RECIPIENT    - where to send the report (defaults to sender)
    SMTP_SERVER         - default smtp.office365.com
    SMTP_PORT           - default 587
"""

import os
import re
import sys
from datetime import datetime

import requests
from openpyxl import load_workbook

GRAPH_SCOPE = "https://graph.microsoft.com/.default"

URL_REGEX = re.compile(r'(https?://[^\s"\'<>]+)')
REQUEST_TIMEOUT = 10  # seconds
MAX_WORKERS = 8


def extract_links_from_excel(path):
    """
    Pull every link out of an Excel workbook. Catches two cases:
      1. Real hyperlinks attached to a cell (Insert > Link in Excel)
      2. Plain text URLs typed directly into a cell
    Returns a list of dicts: {sheet, cell, url}
    """
    wb = load_workbook(path, data_only=True)
    links = []
    seen = set()

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                url = None

                # Case 1: a "real" Excel hyperlink object
                if cell.hyperlink is not None and cell.hyperlink.target:
                    url = cell.hyperlink.target

                # Case 2: a URL typed as plain text in the cell
                elif isinstance(cell.value, str):
                    match = URL_REGEX.search(cell.value)
                    if match:
                        url = match.group(1).rstrip('.,;)')

                if url and url.startswith("http"):
                    key = (sheet.title, url)
                    if key not in seen:
                        seen.add(key)
                        links.append({
                            "sheet": sheet.title,
                            "cell": cell.coordinate,
                            "url": url,
                        })
    return links


def check_link(entry):
    """Check one link, return entry enriched with status info."""
    url = entry["url"]
    headers = {"User-Agent": "Mozilla/5.0 (LinkChecker/1.0)"}
    try:
        resp = requests.head(
            url, allow_redirects=True, timeout=REQUEST_TIMEOUT, headers=headers
        )
        # Some servers don't support HEAD properly (405/403) - fall back to GET
        if resp.status_code in (405, 403, 501):
            resp = requests.get(
                url, allow_redirects=True, timeout=REQUEST_TIMEOUT, headers=headers,
                stream=True,
            )
        entry["status_code"] = resp.status_code
        entry["ok"] = resp.ok
        entry["error"] = None
    except requests.exceptions.RequestException as exc:
        entry["status_code"] = None
        entry["ok"] = False
        entry["error"] = str(exc)[:200]
    return entry


def check_all_links(links):
    from concurrent.futures import ThreadPoolExecutor
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        for result in pool.map(check_link, links):
            results.append(result)
    return results


def build_report_html(results):
    broken = [r for r in results if not r["ok"]]
    working = [r for r in results if r["ok"]]

    date_str = datetime.now().strftime("%A, %B %d, %Y")

    def row(r):
        status = r["status_code"] if r["status_code"] else "No response"
        error = f"<br><small>{r['error']}</small>" if r["error"] else ""
        color = "#1a7f37" if r["ok"] else "#c0392b"
        return f"""
        <tr>
            <td style="padding:6px 10px;border:1px solid #ddd;">{r['sheet']}</td>
            <td style="padding:6px 10px;border:1px solid #ddd;">{r['cell']}</td>
            <td style="padding:6px 10px;border:1px solid #ddd;word-break:break-all;">
                <a href="{r['url']}">{r['url']}</a>
            </td>
            <td style="padding:6px 10px;border:1px solid #ddd;color:{color};font-weight:bold;">
                {status}{error}
            </td>
        </tr>"""

    broken_rows = "".join(row(r) for r in broken) or (
        '<tr><td colspan="4" style="padding:10px;">None 🎉</td></tr>'
    )
    working_rows = "".join(row(r) for r in working) or (
        '<tr><td colspan="4" style="padding:10px;">None</td></tr>'
    )

    summary_color = "#c0392b" if broken else "#1a7f37"
    summary_text = (
        f"{len(broken)} broken link(s) found" if broken else "All links are working"
    )

    html = f"""
    <html>
    <body style="font-family:Arial,Helvetica,sans-serif;color:#222;">
        <h2>Daily Link Check — {date_str}</h2>
        <p style="font-size:16px;color:{summary_color};font-weight:bold;">
            {summary_text} ({len(working)}/{len(results)} OK)
        </p>

        <h3 style="color:#c0392b;">Broken Links</h3>
        <table style="border-collapse:collapse;width:100%;">
            <tr style="background:#f4f4f4;">
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">Sheet</th>
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">Cell</th>
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">URL</th>
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">Status</th>
            </tr>
            {broken_rows}
        </table>

        <h3 style="color:#1a7f37;margin-top:24px;">Working Links</h3>
        <table style="border-collapse:collapse;width:100%;">
            <tr style="background:#f4f4f4;">
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">Sheet</th>
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">Cell</th>
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">URL</th>
                <th style="padding:6px 10px;border:1px solid #ddd;text-align:left;">Status</th>
            </tr>
            {working_rows}
        </table>
    </body>
    </html>
    """
    return html, len(broken)


def get_graph_access_token():
    """Client-credentials (app-only) OAuth2 flow against Azure AD."""
    tenant_id = os.environ["GRAPH_TENANT_ID"]
    client_id = os.environ["GRAPH_CLIENT_ID"]
    client_secret = os.environ["GRAPH_CLIENT_SECRET"]

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": GRAPH_SCOPE,
    }
    resp = requests.post(token_url, data=data, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.json()["access_token"]


def send_email(html_body, subject):
    sender = os.environ["GRAPH_SENDER_EMAIL"]
    recipient = os.environ.get("REPORT_RECIPIENT", sender)

    access_token = get_graph_access_token()

    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": recipient}}],
        },
        "saveToSentItems": "true",
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
    if resp.status_code >= 300:
        raise RuntimeError(
            f"Graph sendMail failed ({resp.status_code}): {resp.text[:500]}"
        )


def short_url_label(url, max_len=60):
    """
    Long query strings (common in scheduling-widget URLs) blow out table
    width and push the Status column off-screen. Show a short label as the
    link text while keeping the full URL as the actual href.
    """
    if len(url) <= max_len:
        return f"[{url}]({url})"
    return f"[{url[:max_len]}…]({url})"


def write_step_summary(results):
    """
    Writes a formatted report to the GitHub Actions run summary page
    (the "Summary" tab you see when you open a workflow run). No-ops when
    not running inside GitHub Actions.
    """
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return

    date_str = datetime.now().strftime("%A, %B %d, %Y")
    total = len(results)
    broken = [r for r in results if not r["ok"]]
    ok = total - len(broken)

    lines = [f"## 🔗 Daily Link Check — {date_str}", ""]

    if total == 0:
        lines.append("No links were found in the workbook.")
    elif broken:
        lines.append(f"### ❌ {len(broken)} broken link(s) found ({ok}/{total} OK)")
        lines.append("")
        lines.append("| Sheet | Cell | URL | Status |")
        lines.append("|---|---|---|---|")
        for r in broken:
            status = r["status_code"] if r["status_code"] else "No response"
            err = f" — {r['error']}" if r["error"] else ""
            lines.append(
                f"| {r['sheet']} | {r['cell']} | {short_url_label(r['url'])} | {status}{err} |"
            )
    else:
        lines.append(f"### ✅ All {total} links working")

    if results:
        lines.append("")
        lines.append("<details><summary>Show all links checked</summary>")
        lines.append("")
        lines.append("| | Sheet | Cell | URL | Status |")
        lines.append("|---|---|---|---|---|")
        for r in results:
            icon = "✅" if r["ok"] else "❌"
            status = r["status_code"] if r["status_code"] else "No response"
            lines.append(
                f"| {icon} | {r['sheet']} | {r['cell']} | {short_url_label(r['url'])} | {status} |"
            )
        lines.append("")
        lines.append("</details>")

    with open(summary_path, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def main():
    excel_path = os.environ.get("EXCEL_FILE_PATH", "links.xlsx")

    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found at '{excel_path}'", file=sys.stderr)
        sys.exit(1)

    print(f"Reading links from {excel_path} ...")
    links = extract_links_from_excel(excel_path)
    print(f"Found {len(links)} link(s). Checking...")

    results = check_all_links(links) if links else []
    html, broken_count = build_report_html(results) if results else (
        "<p>No links were found in the workbook.</p>", 0
    )

    # Write the dashboard summary first, so you see results in the Actions
    # tab even if the email step below fails.
    write_step_summary(results)

    subject = (
        f"⚠️ Link Check: {broken_count} broken link(s)"
        if broken_count
        else "✅ Link Check: all links working"
    )

    send_email(html, subject)
    print(f"Report sent. {broken_count} broken link(s).")


if __name__ == "__main__":
    main()

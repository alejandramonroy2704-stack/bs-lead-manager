"""
BS Weekly Lead Manager — Flask backend con sistema de login
"""
import os, sys, io, re, json, csv
from datetime import datetime
from functools import wraps
from flask import Flask, jsonify, request, render_template, send_file, \
                  session, redirect, url_for
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import database as db

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "bs-lead-manager-secret-2024")


def _load_env():
    """Load .env file from project directory if present."""
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env_path):
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())

_load_env()

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ── Usuarios autorizados ───────────────────────────────────────────────────────
USERS = {
    "bscommercialkitchens": os.environ.get("APP_PASSWORD", "BS2024admin"),
}

# ── Login decorator ────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

# ── Login / Logout routes ──────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("logged_in"):
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "").strip()
        # Reload USERS each time so env var changes take effect
        users = {"bscommercialkitchens": os.environ.get("APP_PASSWORD", "BS2024admin")}
        if username in users and users[username] == password:
            session["logged_in"] = True
            session["username"]  = username
            return redirect(url_for("index"))
        else:
            error = "Incorrect username or password"
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Startup check ─────────────────────────────────────────────────────────────
def check_api_key():
    if not ANTHROPIC_API_KEY:
        print("\n" + "=" * 60)
        print("  WARNING: ANTHROPIC_API_KEY is not set!")
        print("=" * 60 + "\n")


# ── Prompt builder ────────────────────────────────────────────────────────────
def build_prompt(state: str, city: str, category: str, quantity: int, existing: list) -> str:
    city_str = f", {city}" if city.strip() else ""

    if category == "all":
        cat_filter = (
            "Include a mix of categories: restaurants, hotels, casinos, "
            "catering companies, and food trucks. Aim for variety."
        )
    else:
        cat_filter = f"Generate ONLY '{category}' businesses."

    names_to_exclude = [name for name, _ in existing[:150]]
    exclude_block = ""
    if names_to_exclude:
        exclude_block = (
            "\n\nDO NOT include these businesses (already in our database):\n"
            + "\n".join(f"- {n}" for n in names_to_exclude)
        )

    return f"""You are generating sales leads for B+S Commercial Kitchens, a 70-year-old Australian commercial kitchen equipment supplier (woks, fryers, combi ovens, refrigeration, prep benches, dishwashers, bain maries, etc.).

Generate exactly {quantity} unique Australian hospitality business sales leads in {state}{city_str}.
{cat_filter}
{exclude_block}

CRITICAL: Return ONLY a raw JSON array. No markdown fences, no explanation, no extra text. Just the JSON array starting with [ and ending with ].

Each object must have EXACTLY these fields:
{{
  "name": "Business trading name",
  "category": "restaurant" | "hotel" | "casino" | "catering" | "foodtruck",
  "email": "realistic email e.g. info@name.com.au or chef@name.com.au",
  "phone": "Australian phone e.g. 03 9xxx xxxx or 0412 xxx xxx",
  "address": "Street number and name",
  "suburb": "Real suburb in {state}",
  "state": "{state}",
  "website": "https://www.businessname.com.au",
  "priority": "high" | "medium" | "low",
  "reason": "1-2 sentences why they need commercial kitchen equipment from B+S"
}}

Priority rules:
- high = 5-star hotels, major casino dining venues, large corporate catering (50+ staff), multi-location restaurant chains
- medium = independent full-service restaurants, 3-4 star hotels, established catering companies
- low = small cafes, food trucks, small family restaurants, pop-ups

Use real suburb names in {state}. Phone numbers must be realistic Australian format. All websites must end in .com.au or .net.au. Reasons should mention specific equipment types (combi ovens, fryers, bain maries, etc.).
Generate exactly {quantity} objects."""


def parse_leads(text: str) -> list:
    text = text.strip()
    try:
        data = json.loads(text)
        if isinstance(data, list): return data
    except Exception: pass
    match = re.search(r"\[[\s\S]*\]", text)
    if match:
        try:
            data = json.loads(match.group(0))
            if isinstance(data, list): return data
        except Exception: pass
    clean = re.sub(r"```(?:json)?", "", text).strip().strip("`")
    try:
        data = json.loads(clean)
        if isinstance(data, list): return data
    except Exception: pass
    return []


# ── Email template builder ────────────────────────────────────────────────────
def build_email_template(lead: dict) -> dict:
    """Return {"subject": str, "body": str} for the email modal."""
    name     = lead.get("name", "")
    suburb   = lead.get("suburb", "")
    interest = lead.get("interest_product", "")
    reason   = lead.get("reason", "")
    product_line = interest if interest else "commercial kitchen equipment"
    subject  = f"B+S Commercial Kitchens – Quality {product_line.title()} for {name}"
    body = f"""Dear {name} Team,

I hope this message finds you well.

My name is [Your Name] from B+S Commercial Kitchens. We have been supplying Australian hospitality businesses with premium commercial kitchen equipment for over 70 years.

I'm reaching out because we believe {name} in {suburb} could benefit from our range of {product_line}. {reason}

At B+S Commercial Kitchens we offer:
• Energy-efficient cooking equipment (combi ovens, fryers, woks, grills)
• Professional refrigeration and cold rooms
• Prep benches, bain maries, and food holding equipment
• Dishwashers and warewashing solutions
• After-sales service and spare parts Australia-wide

We would love to arrange a brief chat or showroom visit at your convenience — no obligation, just a conversation about how we can support your kitchen.

Please feel free to reply to this email or call us at (03) 9460 5655.

Warm regards,

[Your Name]
B+S Commercial Kitchens
📞 (03) 9460 5655
🌐 www.bscommercialkitchens.com.au
📍 Reservoir, VIC"""
    return {"subject": subject, "body": body}


# ── Excel helpers ─────────────────────────────────────────────────────────────
HEADERS    = ["Name", "Category", "Email", "Phone", "Address",
              "Suburb", "State", "Website", "Priority", "Status",
              "Interest / Product", "Reason", "Notes", "Sent"]
COL_WIDTHS = [35, 14, 32, 18, 28, 18, 8, 36, 10, 12, 22, 52, 30, 8]
ORANGE     = "E85D04"
PRIO_FILL  = {"high": "FFE0E0", "medium": "FFF8E0", "low": "E6F4E6"}
STATUS_FILL = {
    "pending":    "FFFBEB",
    "contacted":  "EFF6FF",
    "interested": "F5F3FF",
    "meeting":    "ECFEFF",
    "closed":     "F0FDF4",
    "lost":       "F9FAFB",
}


def _style_sheet(ws, leads: list):
    ws.append(HEADERS)
    hdr_fill = PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    for i, lead in enumerate(leads, 2):
        prio   = lead.get("priority", "medium")
        status = lead.get("status", "pending")
        row_data = [
            lead.get("name", ""),
            lead.get("category", ""),
            lead.get("email", ""),
            lead.get("phone", ""),
            lead.get("address", ""),
            lead.get("suburb", ""),
            lead.get("state", ""),
            lead.get("website", ""),
            prio.upper(),
            status.upper(),
            lead.get("interest_product", ""),
            lead.get("reason", ""),
            lead.get("notes", ""),
            "Yes" if lead.get("sent") else "No",
        ]
        ws.append(row_data)
        # Priority tint for high-priority rows
        if prio == "high":
            fill_hex = PRIO_FILL["high"]
            row_fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
            for cell in ws[i]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.fill = row_fill
        else:
            # Status tint for other rows
            sfill_hex = STATUS_FILL.get(status, "FFFFFF")
            sfill = PatternFill(start_color=sfill_hex, end_color=sfill_hex, fill_type="solid")
            for cell in ws[i]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.fill = sfill
        ws.row_dimensions[i].height = 28

    for idx, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def _build_workbook_for_week(week_id: int):
    leads = db.get_leads_by_week(week_id)
    wb    = openpyxl.Workbook()
    ws    = wb.active
    ws.title = "Leads"
    _style_sheet(ws, leads)
    return wb, leads


def _build_workbook_all():
    all_leads = db.get_all_leads()
    wb        = openpyxl.Workbook()

    ws_all       = wb.active
    ws_all.title = "All Leads"
    _style_sheet(ws_all, all_leads)

    cats: dict[str, list] = {}
    for lead in all_leads:
        cats.setdefault(lead.get("category", "other"), []).append(lead)

    for cat, cat_leads in sorted(cats.items()):
        ws = wb.create_sheet(title=cat.title()[:31])
        _style_sheet(ws, cat_leads)

    ws_sum = wb.create_sheet(title="Summary")
    ws_sum.append(["Category", "Total", "Contacted+", "Closed", "Lost", "With Email"])
    hdr_fill = PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")
    for cell in ws_sum[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")
    for cat, cat_leads in sorted(cats.items()):
        contacted_plus = sum(1 for l in cat_leads if l.get("status","pending") not in ("pending","lost"))
        closed = sum(1 for l in cat_leads if l.get("status") == "closed")
        lost   = sum(1 for l in cat_leads if l.get("status") == "lost")
        wmail  = sum(1 for l in cat_leads if l.get("email"))
        ws_sum.append([cat.title(), len(cat_leads), contacted_plus, closed, lost, wmail])
    for col in ["A","B","C","D","E","F"]:
        ws_sum.column_dimensions[col].width = 18

    return wb, all_leads


# ── API routes ────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    return render_template("index.html", username=session.get("username", ""))


@app.route("/api/stats")
@login_required
def api_stats():
    return jsonify(db.get_stats())


@app.route("/api/weeks")
@login_required
def api_weeks():
    return jsonify(db.get_weeks())


@app.route("/api/leads")
@login_required
def api_all_leads():
    return jsonify(db.get_all_leads())


@app.route("/api/leads/week/<int:week_id>")
@login_required
def api_week_leads(week_id):
    return jsonify(db.get_leads_by_week(week_id))


@app.route("/api/leads/current")
@login_required
def api_current_leads():
    return jsonify(db.get_current_week_leads())


@app.route("/api/leads/<int:lead_id>/history")
@login_required
def api_lead_history(lead_id):
    return jsonify(db.get_status_history(lead_id))


@app.route("/api/leads/stale")
@login_required
def api_stale_leads():
    days = int(request.args.get("days", 7))
    return jsonify(db.get_stale_leads(days))


@app.route("/api/dashboard/categories")
@login_required
def api_dashboard_categories():
    return jsonify(db.get_pipeline_stats_by_category())


@app.route("/api/dashboard/evolution")
@login_required
def api_dashboard_evolution():
    return jsonify(db.get_weekly_evolution())


@app.route("/api/leads/<int:lead_id>/email-template")
@login_required
def api_email_template(lead_id):
    leads = db.get_all_leads()
    lead  = next((l for l in leads if l["id"] == lead_id), None)
    if not lead:
        return jsonify({"error": "Lead not found"}), 404
    tpl = build_email_template(lead)
    return jsonify({"subject": tpl["subject"], "body": tpl["body"]})


@app.route("/api/generate", methods=["POST"])
@login_required
def api_generate():
    key = os.environ.get("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY)
    if not key:
        return jsonify({"error": "ANTHROPIC_API_KEY not set."}), 400

    data     = request.json or {}
    state    = data.get("state",    "Victoria")
    city     = data.get("city",     "").strip()
    category = data.get("category", "all")
    quantity = min(int(data.get("quantity", 20)), 200)   # max 200

    existing = db.get_all_business_names()
    prompt   = build_prompt(state, city, category, quantity, existing)

    try:
        client  = anthropic.Anthropic(api_key=key)
        message = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=16000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw   = message.content[0].text
        leads = parse_leads(raw)
        if not leads:
            return jsonify({"error": "Could not parse AI response as JSON.",
                            "raw_snippet": raw[:400]}), 500
    except anthropic.AuthenticationError:
        return jsonify({"error": "Invalid Anthropic API key."}), 401
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    week_id = db.get_or_create_current_week()

    # ── Weekly limit check ────────────────────────────────────────────────────
    weekly_limit = db.WEEKLY_LIMIT
    if weekly_limit > 0:
        already   = db.get_weekly_lead_count(week_id)
        remaining = weekly_limit - already
        if remaining <= 0:
            return jsonify({
                "error": f"Weekly limit of {weekly_limit} leads reached. "
                         f"Try again next week or ask your admin to increase WEEKLY_LIMIT."
            }), 429
        if quantity > remaining:
            quantity = remaining
            leads    = leads[:quantity]

    saved = db.save_leads(leads, week_id)

    return jsonify({
        "success":      True,
        "week_id":      week_id,
        "leads":        saved,
        "generated":    len(leads),
        "saved":        len(saved),
        "skipped":      len(leads) - len(saved),
        "weekly_limit": weekly_limit,
        "weekly_count": db.get_weekly_lead_count(week_id),
    })


@app.route("/api/leads/<int:lead_id>/status", methods=["PUT"])
@login_required
def api_update_status(lead_id):
    data             = request.json or {}
    status           = data.get("status", "pending")
    notes            = data.get("notes")
    interest_product = data.get("interest_product")
    db.update_status(lead_id, status, notes, interest_product)
    return jsonify({"success": True})


@app.route("/api/leads/<int:lead_id>/sent", methods=["PUT"])
@login_required
def api_toggle_sent(lead_id):
    sent = (request.json or {}).get("sent", True)
    db.mark_sent(lead_id, sent)
    return jsonify({"success": True})


@app.route("/api/leads/<int:lead_id>", methods=["DELETE"])
@login_required
def api_delete_lead(lead_id):
    db.delete_lead(lead_id)
    return jsonify({"success": True})


# ── Export routes ─────────────────────────────────────────────────────────────

@app.route("/api/export/week/<int:week_id>")
@login_required
def api_export_week(week_id):
    wb, leads = _build_workbook_for_week(week_id)
    if not leads:
        return jsonify({"error": "No leads for this week"}), 404
    week  = db.get_week(week_id)
    label = (week["week_label"].replace(" ", "_").replace("/", "-")
             if week else f"week_{week_id}")
    fname = f"BS_Leads_{label}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    buf   = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/export/all")
@login_required
def api_export_all():
    wb, leads = _build_workbook_all()
    if not leads:
        return jsonify({"error": "No leads in database"}), 404
    fname = f"BS_All_Leads_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    buf   = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/export/csv")
@login_required
def api_export_csv():
    leads = db.get_all_leads()
    if not leads:
        return jsonify({"error": "No leads in database"}), 404
    fields = ["name","category","email","phone","address","suburb","state",
              "website","priority","status","interest_product","reason","notes","sent","week_label"]
    buf = io.StringIO()
    w   = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore")
    w.writeheader()
    for lead in leads:
        lead["sent"] = "Yes" if lead.get("sent") else "No"
        w.writerow({f: lead.get(f, "") for f in fields})
    fname = f"BS_Leads_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    bbuf  = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(bbuf, as_attachment=True, download_name=fname, mimetype="text/csv")


# ── Main ──────────────────────────────────────────────────────────────────────

db.init_db()

if __name__ == "__main__":
    check_api_key()
    os.makedirs("exports", exist_ok=True)
    print("\n" + "=" * 52)
    print("  BS Weekly Lead Manager")
    print("  http://localhost:5000")
    print("=" * 52 + "\n")
    app.run(debug=True, port=5000)

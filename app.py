# ───────────────────────────────────────────────────────────────
# Main route
# ───────────────────────────────────────────────────────────────
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def danger_report(path=''):
    instructor_param = request.args.get("instructor", "").strip()
    print(f"[DEBUG] Request instructor={instructor_param}")
    instructor = None
    table_html = None
    message = None
    if instructor_param and instructor_param in INSTRUCTORS:
        instructor = instructor_param
        table_html = get_table_html(instructor)
    else:
        message = (
            '<div style="padding:40px;text-align:center;background:#fafafa;'
            'border-radius:12px;max-width:700px;margin:40px auto;">'
            '<h2>HEADS UP REPORT</h2>'
            '<p>Select an instructor to view student information.</p>'
            '</div>'
        )
    from zoneinfo import ZoneInfo
    updated_at = datetime.now(ZoneInfo("America/Chicago")).strftime("%Y-%m-%d %H:%M:%S %Z")

    import os
    app_env = os.getenv('APP_ENV', 'production')   # ← Add this line

    return render_template(
        "danger_report.html",
        instructor=instructor or "Select an Instructor",
        instructors=INSTRUCTORS,
        table_html=table_html,
        updated_at=updated_at,
        message=message,
        app_env=app_env   # ← Add this line (very important!)
    )
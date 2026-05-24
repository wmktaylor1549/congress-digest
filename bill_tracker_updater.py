def run():
    log.info("=== Bill Tracker Updater ===")

    if not API_KEY:
        log.error("No API key found. Make sure CONGRESS_API_KEY is set.")
        return

    state    = load_state()
    seen_ids = set(state.get("seen_ids", []))
    last_run = state.get("last_run")

    if last_run:
        from_date = last_run[:10]
    else:
        from_date = (datetime.today() - timedelta(days=DAYS_TO_LOOK_BACK)).strftime("%Y-%m-%d")

    log.info(f"Searching for bills introduced since {from_date}")

    raw_bills   = fetch_bills(from_date)
    new_records = []

    for bill in raw_bills:
        uid = f"{bill.get('type')}{bill.get('number')}-{bill.get('congress', 119)}"
        if uid in seen_ids:
            continue
        if not is_climate_energy_bill(bill):
            continue

        # Only include bills introduced on or after from_date
        intro_date = bill.get("introducedDate", "")
        if intro_date and intro_date < from_date:
            log.info(f"Skipping (old intro date {intro_date}): {bill.get('type')} {bill.get('number')}")
            continue

        log.info(f"Processing: {bill.get('type')} {bill.get('number')} — {bill.get('title', '')[:60]}")
        record = build_bill_record(bill)
        new_records.append(record)
        seen_ids.add(uid)

    if new_records:
        log.info(f"\nFound {len(new_records)} new bill(s). Updating tracker...")
        update_excel(new_records)
    else:
        log.info("No new climate/energy bills found since last run.")

    state["seen_ids"] = list(seen_ids)
    state["last_run"] = datetime.utcnow().isoformat()
    save_state(state)
    log.info("Done.")

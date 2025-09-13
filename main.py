def move_sheet_to_excel_and_clear():
    ws = get_sheet()
    values = ws.get_all_values()
    if not values:
        return 0

    # 1行目はヘッダー、2行目以降がデータ
    rows = values[1:] if len(values) > 1 else []
    if not rows:
        return 0

    # A:日付, B:金額, C:使った人（任意）
    cleaned = []
    for r in rows:
        if len(r) >= 2 and r[0] and r[1]:
            cleaned.append([r[0], r[1], r[2] if len(r) > 2 else ""])

    if not cleaned:
        return 0

    # Excelに追記（無ければ作成）
    create_or_update_excel_append(cleaned)

    # スプレッドシートを初期化
    clear_sheet(ws)

    return len(cleaned)


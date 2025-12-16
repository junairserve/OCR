# SN紐づけツール（完全版 v3）

## v3で追加
- 管理者画面：本体SNの「在庫へ戻す」（修理戻り）をワンクリック化
- EVENT_LOG：ステータス変更履歴（いつ/誰/From→To/備考）

## 画面
- 現場：WebアプリURL
- 管理者：WebアプリURL + `?page=admin`

## シート
- PCB_MASTER: pcb_sn, received_date, status, note, imported_at
- BODY_MASTER: body_sn, model, status, note, updated_at
- LINK_LOG:   timestamp_jst, body_sn, pcb_sn, work_type, operator, note
- EVENT_LOG:  timestamp_jst, kind, sn, from_status, to_status, operator, note

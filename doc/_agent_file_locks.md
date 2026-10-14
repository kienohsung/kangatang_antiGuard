# Agent File Locks (`_agent_file_locks.md`)

> Quy ước phối hợp Multi-Agent: Trước khi chỉnh sửa bất kỳ file nào trong dự án, Agent phải kiểm tra bảng dưới đây và đăng ký khóa file (`LOCKED`). Khi hoàn tất (Step 4 Healthcheck pass), giải phóng khóa (`IDLE`).

| File Path | Agent Owner | Status | Task / Description | Timestamp |
| :--- | :--- | :--- | :--- | :--- |
| `addin/KangatangGuard_Code.vba` | MainAgent | `IDLE` | Released after v3.15.0 AccessVBOM Self-Healing & Zero-Popup verification | 2026-10-14 23:45 |
| `addin/Install_KangatangGuard.ps1` | MainAgent | `IDLE` | Released after v3.15.0 deployment | 2026-10-14 23:45 |
| `addin/Install_Addin.bat` | MainAgent | `IDLE` | Released after v3.15.0 deployment | 2026-10-14 23:45 |
| `addin/Deploy_To_Hub.ps1` | MainAgent | `IDLE` | Released after v3.15.0 deployment | 2026-10-14 23:45 |
| `distribution/Install_Client.ps1` | MainAgent | `IDLE` | Released after v3.15.0 deployment | 2026-10-14 23:45 |

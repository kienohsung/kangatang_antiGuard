# TÃ€I LIá»†U KIáº¾N TRÃšC & THIáº¾T Káº¾ Há»† THá»NG
# TÃ¡i cáº¥u trÃºc Kho Sao lÆ°u Máº«u Virus Táº­p trung (Centralized Virus Quarantine Architecture)
**PhiÃªn báº£n phÃ¡t hÃ nh:** v3.8.5 Production-grade  
**NgÃ y cáº­p nháº­t:** 2026-09-13  
**TÃ¡c giáº£:** Tech Lead kiÃªm Senior Full-stack Architect  

---

## 1. Má»¤C TIÃŠU & Bá»I Cáº¢NH THIáº¾T Káº¾ (ARCHITECTURAL RATIONALE)

### 1.1. Báº¥t cáº­p cá»§a mÃ´ hÃ¬nh sao lÆ°u phÃ¢n tÃ¡n cÅ© (_Backup_Kangatang):
1. **Nguy cÆ¡ an ninh ngÆ°á»i dÃ¹ng má»Ÿ nháº§m (Accidental Execution):**
   MÃ´ hÃ¬nh cÅ© táº¡o thÆ° má»¥c _Backup_Kangatang ngay táº¡i thÆ° má»¥c chá»©a tá»‡p nhiá»…m. NgÆ°á»i dÃ¹ng vÄƒn phÃ²ng tháº¥y tá»‡p sao lÆ°u láº¡ cÃ³ thá»ƒ tÃ² mÃ² báº¥m má»Ÿ, khiáº¿n macro virus kÃ­ch hoáº¡t trá»Ÿ láº¡i náº¿u mÃ¡y Ä‘Ã³ chÆ°a cÃ i Add-in báº£o vá»‡.
2. **Ã” nhiá»…m cáº¥u trÃºc thÆ° má»¥c lÃ m viá»‡c:**
   HÃ ng loáº¡t thÆ° má»¥c _Backup_Kangatang xuáº¥t hiá»‡n ráº£i rÃ¡c trÃªn cÃ¡c á»• Ä‘Ä©a dÃ¹ng chung (File Server LAN) vÃ  mÃ¡y cÃ¡ nhÃ¢n, gÃ¢y khÃ³ khÄƒn cho viá»‡c quáº£n lÃ½ tÃ i liá»‡u.
3. **Thiáº¿u kháº£ nÄƒng kiá»ƒm toÃ¡n & truy váº¿t táº­p trung (Lack of Central Auditability):**
   Quáº£n trá»‹ viÃªn IT khÃ´ng náº¯m Ä‘Æ°á»£c tá»•ng thá»ƒ cÃ¡c vá»¥ lÃ¢y nhiá»…m Ä‘ang diá»…n ra trÃªn toÃ n máº¡ng LAN.

### 1.2. Thiáº¿t káº¿ Kiáº¿n trÃºc má»›i v3.8.5 (Centralized Quarantine Hub):
- **Vá»‹ trÃ­ táº­p trung duy nháº¥t:**
  \\192.168.223.7\file_shared\vietnam\z. ETC\2. Virus backupfile - DO NOT OPEN IT
- **Quy táº¯c Ä‘á»‹nh danh tá»‡p (Naming Convention):**
  <TÃªnTá»‡pGá»‘c>_backup_<TÃªnMÃ¡yTráº¡m>_<yyyyMMdd_HHmmss>.<Pháº§nMá»ŸRá»™ng>
  *VÃ­ dá»¥:* BaoCaoKinhDoanh_backup_CM-GA-MRKIENIT1_20260913_162500.xlsx
  => IT láº­p tá»©c xÃ¡c Ä‘á»‹nh Ä‘Æ°á»£c tá»‡p bá»‹ nhiá»…m báº¯t nguá»“n tá»« mÃ¡y tÃ­nh nÃ o trong máº¡ng ná»™i bá»™.
- **CÆ¡ cháº¿ phÃ²ng thá»§ khi Máº¥t káº¿t ná»‘i LAN (Defensive Offline Fallback):**
  Náº¿u mÃ¡y tráº¡m mang ra ngoÃ i cÃ´ng ty (offline khÃ´ng tháº¥y Hub 223.7), há»‡ thá»‘ng tá»± Ä‘á»™ng chuyá»ƒn hÆ°á»›ng cÃ¡ch ly an toÃ n vÃ o thÆ° má»¥c áº©n cá»§a á»©ng dá»¥ng:
  %APPDATA%\KangatangGuard\Quarantine_Backup\
  Tuyá»‡t Ä‘á»‘i khÃ´ng táº¡o báº¥t ká»³ thÆ° má»¥c nÃ o táº¡i thÆ° má»¥c lÃ m viá»‡c cá»§a ngÆ°á»i dÃ¹ng.
- **Bá»™ lá»c chá»‘ng Ä‘á»‡ quy (Anti-Recursion Exclusion Filter):**
  Táº¥t cáº£ cÃ¡c thÃ nh pháº§n Scanner (Add-in, Background Worker, Standalone Scanner) Ä‘á»u loáº¡i trá»« triá»‡t Ä‘á»ƒ thÆ° má»¥c *Virus backupfile* khá»i luá»“ng duyá»‡t tá»‡p Ä‘á»ƒ khÃ´ng quÃ©t ngÆ°á»£c láº¡i kho lÆ°u trá»¯ máº«u virus.

---

## 2. MA TRáº¬N Äá»I CHIáº¾U THAY Äá»”I (BEFORE & AFTER MATRIX)

| TiÃªu chÃ­ | TrÆ°á»›c khi Sá»­a (v3.8.4) | Sau khi Sá»­a (v3.8.5) |
| :--- | :--- | :--- |
| **Vá»‹ trÃ­ lÆ°u trá»¯** | PhÃ¢n tÃ¡n táº¡i thÆ° má»¥c _Backup_Kangatang cáº¡nh file | ðŸ›¡ï¸ Táº­p trung duy nháº¥t táº¡i Hub LAN: 2. Virus backupfile - DO NOT OPEN IT |
| **Rá»§i ro ngÆ°á»i dÃ¹ng má»Ÿ nháº§m** | Ráº¥t cao | ðŸ›¡ï¸ Triá»‡t tiÃªu hoÃ n toÃ n (khÃ´ng cÃ²n folder cáº¡nh file) |
| **Kháº£ nÄƒng truy váº¿t nguá»“n nhiá»…m** | KhÃ´ng rÃµ mÃ¡y nÃ o nhiá»…m | ðŸ›¡ï¸ Gáº¯n Ä‘á»‹nh danh COMPUTERNAME vÃ o tá»«ng file sao lÆ°u |
| **Xá»­ lÃ½ khi Offline LAN** | Váº«n táº¡o folder táº¡i chá»— | ðŸ›¡ï¸ Chuyá»ƒn hÆ°á»›ng lÆ°u vÃ o %APPDATA%\KangatangGuard\Quarantine_Backup\ |
| **Bá»™ lá»c chá»‘ng quÃ©t kho virus** | Chá»‰ lá»c _Backup_Kangatang | ðŸ›¡ï¸ Lá»c thÃªm *Virus backupfile* chá»‘ng quÃ©t Ä‘á»‡ quy |
| **PhiÃªn báº£n phÃ¡t hÃ nh** | v3.8.4 | **v3.8.5 Production-grade** |
"""locales — 各語言的字串表，一個語言一個檔。

每個檔案只匯出一個 `STRINGS: dict[str, str]`。改譯文不會影響任何邏輯：
程式一律用 key 比對，Excel 一律用 A 欄英文機器鍵比對。改錯最壞情況只是
畫面或 B 欄顯示怪怪的。

key 命名空間：
    gui.*    GUI 介面文字
    acct.*   三表科目（Excel B 欄）
    ratio.*  Data_Ratios 列名與算法說明
    xls.*    Index sheet、Data_Meta 顯示名、Segments 軸名
    err.*    會顯示給使用者看的錯誤訊息

繁體中文（zh_tw.py）是母表，其他語言從它翻出來。找不到 key 時 i18n.t()
會退回繁中，所以繁中必須最完整——`tests/test_i18n.py` 會檢查所有語言的
key 集合完全一致。
"""

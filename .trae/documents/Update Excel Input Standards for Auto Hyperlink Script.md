I will modify `e:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink\测试文件\auto_hyperlink.py` to update the input standards as requested.

**Planned Changes:**

1. **Receive Date Format (`收文日期`)**:

   * Modify `_format_received_date` function.

   * Change the default format to `Year.Month.Day` (e.g., `2026.1.9`) to match the user's requirement "yyyy.mm.dd" and the example provided.

2. **File Name Display (`文件名`)**:

   * Modify `_update_workbook` function.

   * When writing the hyperlink for the "文件名" column, use the filename (basename) as the display text instead of the relative path.

3. **Self-Numbering (`自编号`)**:

   * Add a `CATEGORY_PREFIX_MAP` dictionary to map category names to their Pinyin acronyms (e.g., "其他" -> "QT", "上级文" -> "SJW", "事项通知" -> "SXTZ").

   * Modify `_infer_self_id_pattern` (or create a new function `_generate_self_id`) to:

     * Determine the prefix based on the category label (folder name) using the map.

     * Scan the existing "自编号" column to find the maximum sequence number for the current `PREFIX-YEAR` combination.

     * Generate the new ID as `PREFIX-YEAR-SEQUENCE` (starting from 1).

   * Update `_update_workbook` to use this new logic.

4. **Verification**:

   * I will verify the changes by creating a test file and running the script (or simulating the logic) to check if the generated Excel entry matches the requirements.


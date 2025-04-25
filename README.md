**功能概述（自然语言）**  
1. 用户从本地选择并上传一个 Excel 文件（.xlsx 或 .xls）。  
2. 程序读取第一个工作表中指定的“目标和”单元格（默认 A6），并将其值作为要分配的总和。  
3. 在指定的 28 个空格区域（默认 A1:G4）内，随机选出 9 个不同位置，生成 9 个正整数，保证它们之和等于目标值。  
4. 将这 9 个数填入对应单元格后，保留原有工作簿结构和元数据，生成并下载新文件 `processed.xlsx`。  
5. 同时在网页中以 HTML 表格形式预览填充结果。  

---

**Excel 函数及命令（中英双语）**  

| 功能                                    | Excel 函数 / 命令                  | 说明（中）                                    | Explanation (EN)                                        |
|-----------------------------------------|-----------------------------------|-----------------------------------------------|---------------------------------------------------------|
| 随机生成 0–1 之间的数                   | `RAND()`                          | 返回 0 到 1 之间的随机小数                    | Returns a random decimal between 0 and 1                |
| 在指定范围内生成随机整数               | `RANDBETWEEN(bottom, top)`        | 在 bottom 与 top 之间生成随机整数             | Returns a random integer between bottom and top         |
| 求一系列单元格的和                     | `SUM(range)`                      | 计算 range 中所有单元格之和                   | Calculates the sum of the specified range               |
| 累计计数（判断是否已选某位置）         | `COUNTIF(range, criteria)`        | 统计 range 中满足 criteria 的单元格数量       | Counts the number of cells within a range meeting criteria |
| 生成连续整数列表（辅助生成随机位置索引）| `ROW(INDIRECT("1:28"))` 或者 `{1;2;…;28}` | 返回 1 到 28 的连续行号数组                   | Generates an array of integers from 1 to 28             |
| 从列表中抽样（模拟随机抽取位置索引）   | `INDEX(array, RANDBETWEEN(1, n))` | 从 array 中随机抽取一个元素                   | Returns a random element from array using INDEX         |
| 确保单元格中显示整数                   | 设置单元格格式 → 数字 → 0 位小数  | 强制以整数形式显示                            | Format Cells → Number → 0 decimal places                |
| 防止重计算时改变结果                   | 复制 → “选择性粘贴” → 值           | 将公式结果粘贴为静态数值，避免后续再随机         | Copy → Paste Special → Values to freeze the results     |

> **示例组合公式（英文版）**  
> - `=SUM(A1:A9)`：把 A1:A9 的值加总。  
> - `=RANDBETWEEN(1, 10)`：随机生成 1 到 10 的整数。  
> - `=INDEX($A$1:$A$28, RANDBETWEEN(1,28))`：从 A1:A28 中随机选一个位置的值。  

> **示例组合公式（中文版）**  
> - `=SUM(A1:A9)`：计算 A1 到 A9 的总和。  
> - `=RANDBETWEEN(1, 10)`：在 1 到 10 之间生成随机整数。  
> - `=INDEX($A$1:$A$28, RANDBETWEEN(1,28))`：从 A1 至 A28 区域中随机取一格的数值。

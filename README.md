# Duplicate Office.js workbook bug reproduction

Reproduction for bug #####.

To reproduce:

1. Side-load the addin in Desktop Excel (`npm run start:desktop` etc.)
2. Open `repro-workbook.xlsx`
3. Click the Open Task Pane addin button.
4. In the taskpane, click Duplicate Workbook.
5. A copy of the workbook should appear.

In the new workbook copy, the cell containing `=CONTOSO.LOG(Sheet1!A1)` will reload on startup, while the cell containing `=CONTOSO.LOG(A1)` does not. You can also view the addin's console log to confirm that the `=CONTOSO.LOG(Sheet1!A1)` cell is reloaded.

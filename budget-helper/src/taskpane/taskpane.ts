
import { getExpenseList, getRollovers } from '../lookups';
import { preventDefaults, handleFileDrop } from '../file-drop';
import { resetRollover } from '../rollover';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    let dropArea = document.getElementById('drop-area');

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
      dropArea.addEventListener(eventName, preventDefaults, false)
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dropArea.addEventListener(eventName, (e) => {
        dropArea.classList.add('highlight');
      }, false)
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropArea.addEventListener(eventName, (e) => {
        dropArea.classList.remove('highlight');
      }, false)
    });

    dropArea.addEventListener('drop', handleFileDrop, false);

    document.getElementById("reset").onclick = TriggerResetRollovers;

    window.onload = async () => {
      await populateExpenseDropdown();
    };
  }
});

export async function TriggerResetRollovers() {
  try {

    var month = parseInt((<HTMLInputElement>document.getElementById("month-input")).value);
    var year = parseInt((<HTMLInputElement>document.getElementById("year-input")).value);

    let selectElement = document.getElementById('expense-dropdown') as HTMLSelectElement;
    let selectedOption = selectElement.options[selectElement.selectedIndex];
    let selectedExpense = selectedOption.text;

    if (selectedExpense === "All Expenses") {
      selectedExpense = null;
    }

    await Excel.run(async (context) => {
      debugger;

      await resetRollover(month, year, selectedExpense);

      const range = context.workbook.getSelectedRange();
      range.values = [["Done"]];

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function populateExpenseDropdown() {
  try {
    const expenseList = await getExpenseList();
    const expenseDropdown = document.getElementById('expense-dropdown') as HTMLSelectElement;

    // Ensure the dropdown is clear before adding new options
    expenseDropdown.innerHTML = '<option value="">All Expenses</option>';

    for (const expense of expenseList) {
      const option = document.createElement('option');
      option.value = option.text = expense;
      expenseDropdown.add(option);
    }
  } catch (error) {
    console.error('Could not populate the expense dropdown:', error);
  }
}


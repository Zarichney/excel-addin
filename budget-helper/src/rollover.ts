
import { filter, forkJoin, from, map, of, reduce, switchMap, tap, toArray } from "rxjs";
import { AddToTable, TableRows$, UpdateTableRow } from "./excel-helpers";
import { getBudget, getExpenseList, getInitialAmount, getTransactions } from "./lookups";
import { Transaction } from "./transaction";


export class RolloverEntry {
    Month: number;
    Year: number;
    Expense: string;
    Expenses: number;
    BOM: number;
    EOM: number;
}

export async function getRollover(month: number, year: number, expense: string): Promise<RolloverEntry> {
    return new Promise((resolve, reject) => {
        TableRows$('Rollovers').pipe(
            filter((entry: RolloverEntry) => entry.Month === month && entry.Year === year),
            filter((entry: RolloverEntry) => entry.Expense === expense),
            toArray(),
            switchMap(rows => {

                if (rows.length > 1) {
                    console.warn("Unexpected multiple rollover entries found for the same month, year, and expense", rows);
                    debugger;
                } else if (rows.length === 1) {
                    // Match found, continue
                    return of(rows[0]);
                }

                // No results, add in new entry instead
                // First retrieve the total expenses for the month
                const transactions$ = TableRows$('Transactions').pipe(
                    filter((transaction: Transaction) => transaction.Month === month && transaction.Year === year),
                    filter((transaction: Transaction) => transaction.Expense === expense),
                    reduce((acc, transaction) => { return acc + transaction.Amount; }, 0)
                );

                // And the expense's initial amount
                const initialAmount$ = from(getInitialAmount(expense));

                const budget$ = from(getBudget(expense, month, year));

                return forkJoin([transactions$, initialAmount$, budget$]).pipe(
                    switchMap(([monthlyExpenses, initialAmount, budget]) => {
                        const newEntry: RolloverEntry = {
                            Month: month,
                            Year: year,
                            Expense: expense,
                            Expenses: monthlyExpenses,
                            BOM: initialAmount,
                            EOM: initialAmount
                        };

                        // Convert the Promise returned by AddToTable into an Observable
                        return from(AddToTable('Rollovers', newEntry)).pipe(map(() => newEntry));
                    })
                );
            }),
            tap(entry => resolve(entry))
        ).subscribe({
            error: (err) => reject(err),
        });
    });
}

export async function updateRollover(entry: RolloverEntry): Promise<void> {
    return new Promise((resolve, reject) => {
        TableRows$('Rollovers').pipe(
            toArray(), // Collect all rows into an array
            switchMap((rows) => {
                // Find the row that matches the criteria
                const rowIndex = rows.findIndex(row =>
                    row.Month === entry.Month &&
                    row.Year === entry.Year &&
                    row.Expense === entry.Expense
                );

                if (rowIndex === -1) {
                    // No matching row found, consider how you want to handle this
                    console.error("No matching row found to update.");
                    reject(new Error("No matching row found."));
                    return from([null]); // Just to fit the switchMap expected return
                } else {
                    // Update the row. Assuming UpdateTableRow function exists and works similarly to WriteToTable,
                    // but you might need to implement it based on how your Excel integration is set up
                    return from(UpdateTableRow('Rollovers', rowIndex, entry));
                }
            }),
            tap({
                next: () => resolve(),
                error: (err) => reject(err),
            })
        ).subscribe();
    });
}

export async function resetRollover(startingMonth: number, startingYear: number, expense: string | null = null): Promise<void> {

    let expenses: string[];

    if (expense) {
        expenses = [expense];
    } else {
        expenses = await getExpenseList();
    }

    let today = new Date();
    let todaysMonth = today.getMonth() + 1;
    let todaysYear = today.getFullYear();

    for (const expense of expenses) {

        let loopLimit = 24; // Hard max of 2 years of updates

        let month = startingMonth;
        let year = startingYear;

        // validation check to make sure dates are not in the future
        if (year > todaysYear || (year === todaysYear && month > todaysMonth)) {
            console.error("Cannot reset rollover for a future date.");
            return;
        }

        while ((year < todaysYear || (year === todaysYear && month <= todaysMonth)) && loopLimit > 0) {

            const rollover = await getRollover(month, year, expense);

            const budget = await getBudget(expense, month, year);

            let previousMonth = month - 1;
            let previousYear = year;
            if (previousMonth === 0) {
                previousMonth = 12;
                previousYear--;
            }

            const previousRollover = await getRollover(previousMonth, previousYear, expense);

            const monthlyExpenses = await getTransactions(month, year, expense);

            const totalAmount = monthlyExpenses.reduce((total, transaction) => total + transaction.Amount, 0);
            debugger;
            rollover.Expenses = totalAmount;
            rollover.BOM = previousRollover.EOM;
            rollover.EOM = rollover.BOM + budget + totalAmount;

            await updateRollover(rollover);

            month++;
            if (month === 13) {
                month = 1;
                year++;
            }
            loopLimit--;
        }
    }
}

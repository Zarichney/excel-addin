
import { filter, map, of, reduce, switchMap, tap, toArray } from 'rxjs';
import { NamedRangeValues$, TableRows$ } from './excel-helpers';
import { RolloverEntry } from './rollover';
import { Transaction } from './transaction';
import * as rollover from './rollover';

export async function getAccounts(): Promise<{ [key: string]: string }> {
    // returns collection of key value pair
    return new Promise((resolve, reject) => {
        TableRows$('Accounts').pipe(
            map(row => {
                var obj = {};
                obj[row['Number']] = row['Name'];
                return obj;
            }),
            reduce((acc, obj) => ({ ...acc, ...obj }), {})
        ).subscribe({
            next: (obj) => resolve(obj),
            error: (err) => reject(err),
        });
    });
}

export class MatchSet {
    'Match 1': string;
    'Match 2': string;
    'Amount': string;
    'Expense Type': string;
}

export async function getMatchingRules(): Promise<MatchSet[]> {
    return new Promise((resolve, reject) => {
        TableRows$('MatchingRules').pipe(
            toArray()
        ).subscribe({
            next: (rows) => resolve(rows),
            error: (err) => reject(err),
        });
    });
}

export async function getExpenseList(): Promise<string[]> {
    return new Promise((resolve, reject) => {
        NamedRangeValues$('Expenses').pipe(
            toArray(),
            tap(values => resolve(values))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getTransactions(month: number, year: number, expense: string): Promise<Transaction[]> {
    return new Promise((resolve, reject) => {
        TableRows$('Transactions').pipe(
            // tap((transaction: Transaction) => console.log(transaction)),
            filter((transaction: Transaction) => transaction.Month === month && transaction.Year === year),
            filter((transaction: Transaction) => transaction['Expense Type'] === expense),
            toArray(),
            tap(transactions => resolve(transactions))
        ).subscribe({
            error: (err) => reject(err),
        });
    });

}

export async function getAllTransactionIds(): Promise<string[]> {
    return new Promise((resolve, reject) => {
        NamedRangeValues$('TransactionIds').pipe(
            toArray(),
            tap(values => resolve(values))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getRollovers(): Promise<RolloverEntry[]> {
    return new Promise((resolve, reject) => {
        TableRows$('Rollovers').pipe(
            toArray()
        ).subscribe({
            next: (rows) => resolve(rows),
            error: (err) => reject(err),
        });
    });
}

export async function getInitialAmount(expense: string): Promise<number> {
    return new Promise((resolve, reject) => {
        TableRows$('ExpenseData').pipe(
            filter(row => row['Expense Type'] === expense),
            map(row => row['Init']),
            toArray(),
            tap(values => resolve(parseFloat(values[0])))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getBudget(expense: string, month: number | null = null, year: number | null = null): Promise<number> {
    return new Promise((resolve, reject) => {

        const currentBudget$ = TableRows$('ExpenseData').pipe(
            filter(row => row['Expense Type'] === expense),
            map(row => row.Budget)
        );

        let dataSource$ = currentBudget$;

        if (month && year) {
            // Request for specific month and year, look first into change history
            dataSource$ = TableRows$('BudgetHistory').pipe(
                filter(row => row['Expense'] === expense),
                filter(row => (row['Month Start'] <= month && month <= row['Month End'])
                    && (row['Year Start'] <= year && year <= row['Year End'])
                ),
                toArray(),
                switchMap(rows => {

                    if (rows.length > 1) {
                        console.warn("Unexpected multiple rollover entries found for the same month, year, and expense", rows);
                        debugger;
                    } else if (rows.length === 1) {
                        // Match found, continue
                        return of(rows[0].Amount);
                    }

                    // If no historical entry are found, resume fetching from current budget
                    return currentBudget$;
                })
            );
        }

        dataSource$.pipe(
            toArray(),
            tap(values => resolve(parseFloat(values[0])))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getRollover(month: number, year: number, expense: string): Promise<RolloverEntry> {
    return rollover.getRollover(month, year, expense);
}

import { WriteToTable } from "./excel-helpers";
import { ProcessTransactions } from "./transaction";

export function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

export function handleDragOver(event) {
    event.stopPropagation();
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy'; // Explicitly show this is a copy.
}

export function handleFileDrop(event) {

    var files = event.dataTransfer.files;
    if (files.length > 0) {
        var file = files[0];
        var reader = new FileReader();

        reader.onload = ProcessFileDrop;

        reader.readAsText(file);
    }
}

export async function ProcessFileDrop(event) {
    console.log('File upload event:', event);
    var contents = event.target.result;

    const transactions = await ProcessTransactions(contents);

    await WriteToTable("Transactions", transactions);
}
function confirmation() {
    if (confirm("Are you sure you want to delete?"))
        return true;
    else return false;
}

function display(message) {
    alert(message);
}

function warning() {
    alert('WARNING: Changing the Data Source may lose your unsaved changes.');
}
function confirmation() {
    if (confirm("Are you sure you want to delete?"))
        return true;
    else return false;
}

function success() {
    alert("You successfully saved this Entry.");
}
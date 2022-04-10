// Listen for submit

document.getElementById("loan-form").addEventListener("submit", computeResults);

function computeResults(e) {
    // UI

    const UIamount = document.getElementById("amount").value;
    const UIinterest = document.getElementById("interest").value;

    // Calculate

    const principal = parseFloat(UIamount);
    const bal = parseFloat(UIinterest);


    //Compute Total Payment

    const totalPayment = (principal - 1500).toFixed(2);
    let total;
    let percent;
    //Show results
    if (bal >= 51 && bal < 61) {
        total = totalPayment * 0.1;
        percent = 10;
    }
    else if (bal >= 61 && bal < 71) {
        total = totalPayment * 0.2;
        percent = 20;
    }
    else if (bal >= 71 && bal < 81) {
        total = totalPayment * 0.3;
        percent = 30;
    }
    else if (bal >= 81 && bal < 91) {
        total = totalPayment * 0.7;
        percent = 70;
    }
    else if (bal >= 91 && bal <= 100) {
        total = totalPayment * 1;
        percent = 100;
    }
    else {
        total = 0;
        percent = 0;
    }

    var result = UIamount - total;
    document.getElementById("totalPayment").innerHTML = "₼" + parseFloat(total);
    document.getElementById("totalInterest").innerHTML = "%" + percent;
    document.getElementById("monthlyPayment").innerHTML = "₼" + result;

    e.preventDefault();
}

function percent() {
    var res = document.getElementById("percentage").value;
    per = res * 100 / 70;

    document.getElementById("interest").value = per.toFixed(2);
    console.log(document.getElementById("interest").value);
}

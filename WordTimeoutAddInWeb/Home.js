"use strict";
const LIMIT = 10000;
const PAUSE = 0;
var body, landingMain, progressMain, progressFooter, tableMain;
var btnStart, prgText, tblText;

Office.initialize = function (reason) {
    body = document.getElementById("body");
    landingMain = document.getElementById("landingMain");
    progressMain = document.getElementById("progressMain");
    progressFooter = document.getElementById("progressFooter");
    tableMain = document.getElementById("tableMain");
    btnStart = new fabric["Button"](document.getElementById("btnStart"));
};

function process() {
    Word.run(function (context) {
        var i = 0;

        landingMain.style.display = "none";
        body.classList.toggle("ms-notification-progress-determinate");
        progressMain.style.display = "flex";
        progressFooter.style.display = "inline-flex";
        prgText = new fabric["ProgressIndicator"](document.getElementById("prgText"));
        tblText = new fabric["Table"](document.getElementById("tblText"));

        prgText.setTotal(LIMIT);
        setTimeout(writing, PAUSE);
        return context.sync();

        function writing() {
            var row = tblText.container.insertRow(i + 1);
            row.insertCell(0).innerHTML = "text";
            row.insertCell(1).innerHTML = i;
            i++;
            prgText.setProgress(i);
            if (i < LIMIT) {
                setTimeout(writing, PAUSE);
            } else {
                progressMain.style.display = "none";
                progressFooter.style.display = "none";
                body.classList.toggle("ms-font-l");
                body.classList.add("ms-landing-page");
                tableMain.style.display = "flex";
            }
        }
    })
        .catch(function (error) { console.log(error); });
}
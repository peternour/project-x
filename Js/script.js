"use strict";
const calcBtn = document.querySelector(".calc-btn");
const calcLbl = document.querySelector(".date-lbl");
const dateIn = document.getElementById("input");
const header = document.querySelector("h1");

calcBtn.addEventListener("click", function () {
 // Parse the input date
 let birthDate = new Date(dateIn.value);

 let today = new Date();
 // Calculate the difference in years
 let years = today.getFullYear() - birthDate.getFullYear();

 // Adjust the year if the birthday hasn't occurred yet this year
 let months = today.getMonth() - birthDate.getMonth();
 if (months < 0 || (months === 0 && today.getDate() < birthDate.getDate())) {
  years--;
  months += 12;
 }
 let days = today.getDate() - birthDate.getDate();
 if (days < 0) {
  months--;
  // Get the number of days in the previous month
  let prevMonth = new Date(today.getFullYear(), today.getMonth(), 0).getDate();
  days += prevMonth;
 }

 if (years < 21) {
  calcLbl.textContent = `You are UNDER_AGE 21 You are ${years} years, ${months} months, and ${days} days old. And Need to Wait ⌛ ${
   21 - years
  } Years To Come Here 😋`;
 } else if (years > 60) {
  calcLbl.textContent = `👳‍♂️ You are OVER_AGE 60 You are ${years} years, ${months} months, and ${days} days old And Above 60 By ${
   years - 60
  } Year.`;
 } else {
  calcLbl.textContent = `✔ You are ${years} years, ${months} months, and ${days} days old.`;
 }
});

header.addEventListener("click", function () {
 header.textContent = "Peter Nour NagaHamadi Credit Unit";
});

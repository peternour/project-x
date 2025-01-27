function calculateAge() {
  // Prompt the user to input their birthdate in the format YYYY-MM-DD
  let birthDateInput = prompt("Enter your birthdate (YYYY-MM-DD):");
  
  // Parse the input date
  let birthDate = new Date(birthDateInput);
  let today = new Date();

  // Calculate the difference in years
  let years = today.getFullYear() - birthDate.getFullYear();

  // Adjust the year if the birthday hasn't occurred yet this year
  let months = today.getMonth() - birthDate.getMonth();
  if (months < 0 || (months === 0 && today.getDate() < birthDate.getDate())) {
    years--;
    months += 12;
  }

  // Calculate the difference in days
  let days = today.getDate() - birthDate.getDate();
  if (days < 0) {
    months--;
    // Get the number of days in the previous month
    let prevMonth = new Date(today.getFullYear(), today.getMonth(), 0).getDate();
    days += prevMonth;
  }

  // Show the age in an alert
  alert(`You are ${years} years, ${months} months, and ${days} days old.`);
}

calculateAge();

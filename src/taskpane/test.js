function calculateEVM(planEffort, planCompleted, actualEffort, actualCompleted) {
  var PV = planEffort * planCompleted;
  var EV = planEffort * actualCompleted;
  var AC = actualEffort;

  var CPI = EV / AC;
  var SPI = EV / PV;

  return { CPI: CPI, SPI: SPI };
}

// Example usage:
var planEffort = 100; // Planned effort (e.g., hours)
var planCompleted = 0.5; // Percentage of planned work completed (e.g., 50%)
var actualEffort = 120; // Actual effort (e.g., hours)
var actualCompleted = 0.4; // Percentage of actual work completed (e.g., 40%)

var evm = calculateEVM(planEffort, planCompleted, actualEffort, actualCompleted);
console.log("Cost Performance Index (CPI):", evm.CPI);
console.log("Schedule Performance Index (SPI):", evm.SPI);

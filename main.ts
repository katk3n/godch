import { DetailsSheet } from "./details";
import { PaymentsSheet } from "./payments";
import { BalanceSheet } from "./balance";

function createSheets() {
  try {
    const details = new DetailsSheet();
    details.initialize();
    const payments = new PaymentsSheet();
    payments.initialize();
    const balance = new BalanceSheet();
    balance.initialize();
  } catch (e) {
    console.error(e);
    Browser.msgBox(e);
  }
}

function clearSheets() {
  const details = new DetailsSheet();
  details.clear();
  const payments = new PaymentsSheet();
  payments.clear();
  const balance = new BalanceSheet();
  balance.clear();
}

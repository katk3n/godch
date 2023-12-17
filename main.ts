import { DetailsSheet } from "./details";
import { PaymentsSheet } from "./payments";

function createSheets() {
  try {
    const details = new DetailsSheet();
    details.initialize();
    const payments = new PaymentsSheet();
    payments.initialize();
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
}

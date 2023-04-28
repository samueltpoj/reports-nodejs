import { formatExcel } from "./functions/formatExcel.js";
import { sendEmail } from "./functions/sendMail.js";

await formatExcel();
await sendEmail();

# Changes planned for the old system
## Interface Fixes
 - Buttons
 - Windows 64 Pointer
 - File Directory
## Search Fixes
 - Directory Fixes
 - Error Handling
## New System
### Interface V2
 - All functionality of the original system but programed from the ground up to be safer and more efficient.
 - New Systems - Convert button, etc.

# Current System
Got it. I’ve gone through the interface you shared and will break down the functionality into a detailed list, grouped by section.

---

### **1. Filters / Status (Top Left)**

Checkbox options to filter jobs by stage:

* **New Enquiries** – Show only new customer enquiries.
* **Quotes to be Submitted** – Show jobs awaiting quote submission.
* **WIP to be Sequenced** – Show work-in-progress items waiting for scheduling.
* **Jobs In WIP** – Show all active work-in-progress jobs.

---

### **2. Enquiry / Job List Panel (Left Center)**

* Displays a **list of enquiries, quotes, or jobs** depending on filters selected.
* Allows selection of an item to preview or edit.

---

### **3. Counts (Left Bottom)**

* **Enquiries: \*** – Shows the total count of enquiries.
* **Quotes: \*** – Shows the total count of quotes.
* **WIP: \*** – Shows the total count of work-in-progress jobs.

---

### **4. Quick Action Buttons (Left Bottom)**

* **Contract Work** – Opens or manages contract-based jobs.
* **WIP Report** – Generates a work-in-progress report.
* **Jump The Gun** – Likely an urgent/priority job handling option.
* **Show Contracts Folder** – Opens the contracts folder (probably a filesystem location).

---

### **5. Job/Enquiry Preview Panel (Middle)**

Editable job details section:

* **Customer** – Customer name.
* **Contact** – Contact person.
* **Code** – Internal reference code.
* **Grade** – Material/quality grade.
* **Description** – Job description.
* **Qty** – Quantity required.
* **Price** – Quoted or agreed price.
* **Comments (on JC)** – Comments that appear on the Job Card.
* **Comments (not on JC)** – Internal comments (not printed on job card).
* **Drw/Sample #** – Drawing or sample reference number.
* **Status** – Job status (e.g., pending, in progress, complete).
* **Enq #** – Enquiry number.
* **Enq Date** – Enquiry date.
* **Quote #** – Quote number.
* **Job #** – Job number.
* **Job Start Date** – When the job began.
* **Lead Time** – Expected completion duration.
* **Inv #** – Invoice number.
* **File Name** – Associated file path or document name.

---

### **6. Action Buttons (Right Side)**

#### **Enquiry Actions**

* **Add Enquiry** – Create a new enquiry.
* **Convert to Quote (Kevin)** – Convert an enquiry into a formal quote (possibly assigned to “Kevin”).

#### **Quote Actions**

* **Quote Submitted** – Mark quote as submitted to customer.
* **Accept Quote** – Mark quote as accepted.

#### **WIP Actions**

* **Open Job (Kevin)** – Open a job in Kevin’s workflow.
* **Close Job** – Mark job as complete.

#### **Job Management**

* **Print JC** – Print Job Card.
* **Search** – Perform job search.
* **Edit WIP File** – Edit details of WIP file.
* **Edit Job Card** – Edit the job card template/details.
* **Create CT Item** – Create a cost tracking (CT) item.
* **Edit CT Item** – Edit cost tracking item.
* **Sort Search** – Sort search results.
* **Edit Search File** – Edit saved search settings.
* **Job History** – View history of job actions/events.
* **Quote History** – View history of quotes.

---

### **7. File Path (Bottom Right)**

* Shows the directory path where related files are stored (currently `/Users/michaeltomlinson/Documents/GitHub/p...`).

---

✅ **In summary:**
This interface is a full **Job & Quote Management System**. It covers the workflow from **enquiry → quote → acceptance → WIP → completion → invoicing**, with options for reports, contract work, and history tracking.

Would you like me to **organize this into a structured table (like a functional spec)** so you can use it as documentation?

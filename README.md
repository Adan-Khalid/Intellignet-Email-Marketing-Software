# Intellignet-Email-Marketing-Software
A desktop Email Marketing software designed in python.


# 📧 Bulk Email Sender with Automated Follow-ups

This is a powerful, desktop-based email marketing application built with Python and CustomTkinter. It provides a comprehensive solution for managing and running bulk email campaigns, complete with automated follow-ups, reply tracking, and detailed analytics.

## ✨ Key Features

* **Multi-SMTP Management**: Add, edit, and test multiple SMTP accounts (e.g., Gmail, Outlook) to distribute email sending.
* **Template Management**: Create and manage both initial email templates (HTML or text) and sequential follow-up templates.
* **Campaign Scheduling**: Launch campaigns immediately or schedule them to run at a specific date and time.
* **Automated Follow-ups**: Automatically sends follow-up emails to recipients who have not replied after a set period.
* **Intelligent Reply Tracking**:
    * Monitors IMAP inboxes for replies.
    * Automatically stops follow-ups to recipients who have replied.
    * Sends a notification to the admin email for new leads.
* **Notification Center**: See a clean, time-stamped log of all detected replies.
* **DNC Management (Do-Not-Contact)**:
    * Maintain a global **Blocklist** to permanently stop sending to specific emails.
    * Maintain a **Leads List** to stop campaigns for recipients who have become leads (e.g., replied positively).
* **Detailed Analytics**:
    * View high-level stats for all campaigns (Sent, Failed, Date).
    * Drill down into any campaign to see a detailed, email-by-email log.
    * Export campaign logs to CSV.
* **Campaign Resumption**: If a campaign is stopped, the app will offer to resume it from where it left off.
* **Persistent Data**: All SMTP accounts, templates, and campaign logs are saved locally in JSON and log files.

## 🚀 Getting Started

Follow these steps to get the application running on your local machine.

### 1. Prerequisites

* Python 3.8 or newer
* An internet connection
* App Passwords for your email accounts (if using Gmail/Google Workspace). **Do not use your regular password.**

### 2. Installation & Setup

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/your-username/your-repo-name.git](https://github.com/your-username/your-repo-name.git)
    cd your-repo-name
    ```

2.  **Install dependencies:**
    This project uses several Python libraries. Install them using the `requirements.txt` file.
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configure the Application:**
    This is the most important step. Open the `config.py` file in a text editor.
    
    * **Set `ADMIN_EMAIL`**: Change `'your-admin-email@example.com'` to the email address where you want to receive new lead notifications. This is critical for reply tracking.

4.  **Run the Application:**
    ```bash
    python app.py
    ```

### 3. How to Use

1.  **Add SMTP Accounts**: Go to the **SMTP** tab. Add at least one email account using its "App Password". Use the "Test Connection" button to ensure it works.
2.  **Add Templates**: Go to the **Templates** tab.
    * Add at least one subject line.
    * Add at least one **Email Body** (this is for the first email).
    * Add at least one **Follow-up Body** (this is for the automated follow-ups).
3.  **Start a Campaign**: Go to the **Campaign** tab.
    * Give your campaign a name.
    * Upload your list of recipients (a `.txt` or `.csv` file with an 'email' column).
    * Set the delay and click "Start Campaign" or "Schedule Campaign".
4.  **Send Follow-ups**: Go to the **Follow-up** tab. Select a completed campaign and click "Send Follow-ups" to start the process of checking for replies and sending follow-ups to those who haven't.

## 📁 Project File Structure

The application will automatically generate the following files and directories in its root folder:

* `app.py`: The main application source code.
* `config.py`: The central configuration file.
* `requirements.txt`: List of Python dependencies.
* `/bodies/`: A directory where all your `.html` or `.txt` email templates are stored.
* `/logs/`: A directory where detailed JSON logs for every campaign are saved.
* `smtp_list.json`: Stores your SMTP account details.
* `subjects.json`: Stores your list of subject lines.
* `bodies.json`: Stores metadata for your main email templates.
* `followup_bodies.json`: Stores metadata for your follow-up templates.
* `blacklist.json`: Stores your DNC (Lead and Blocklist) emails.
* `notifications.json`: Stores the log for the Notification Center.

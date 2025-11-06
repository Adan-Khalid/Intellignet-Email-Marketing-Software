import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import json
import csv
import time
import threading
import random
import uuid
import re
import imaplib
import email
from collections import defaultdict
import datetime

# --- NEW: Import settings from the config file ---
import config

# Set CustomTkinter theme and color from config
ctk.set_appearance_mode(config.DEFAULT_APPEARANCE_MODE)
ctk.set_default_color_theme(config.DEFAULT_COLOR_THEME)

# ------------------------- 2. Initialization Section ------------------------ #
# Use file/dir names from the config file
files_to_check = [
    config.SMTP_FILE, config.SUBJECTS_FILE, config.EMAIL_BODIES_FILE,
    config.FOLLOWUP_BODIES_FILE, config.BLACKLIST_FILE, config.NOTIFICATIONS_FILE
]

for file in files_to_check:
    if not os.path.exists(file):
        with open(file, 'w') as f:
            # Initialize blacklist as a dictionary
            if file == config.BLACKLIST_FILE:
                json.dump({}, f)
            else:
                json.dump([], f)

for directory in [config.LOG_DIR, config.BODIES_DIR]:
    if not os.path.exists(directory):
        os.makedirs(directory)

# ------------------------- 3. Main Application Class ------------------------ #
class EmailApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # --- Generalized Title ---
        self.title("Bulk Email Sender") 
        self.geometry("1100x750") 
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.viewing_campaign_details_id = None
        
        # Campaign State
        self.running = False
        self.campaign_thread = None
        self.active_campaign_info = {}
        
        # Follow-up State
        self.followup_running = False
        self.followup_thread = None
        self.active_followup_info = {}
        
        self.navigation_frame = None
        self.content_frame = None
        self.all_campaign_logs = {}  # In-memory dictionary for campaign logs
        self.status_var = tk.StringVar(value="Status: Ready") # Initial status
        
        self.scheduled_campaigns = []
        
        # In-memory caches for frequently accessed data
        self.smtp_cache = None
        self.subjects_cache = None
        self.bodies_cache = None
        self.followup_bodies_cache = None
        self.blacklist_cache = None
        self.notifications_cache = None
        
        self.new_notifications_count = tk.IntVar(value=0)
        
        # UI Widget References
        self.progress_bar = None
        self.followup_progress_bar = None
        self.analytics_tree = None

        self.setup_ui()
        # Start a thread to load logs asynchronously at startup
        threading.Thread(target=self._load_initial_data_async, daemon=True).start()
        self._check_schedule()
        threading.Thread(target=self._reply_checker_loop, daemon=True).start()

    def _check_schedule(self):
        """Checks every minute if a scheduled campaign is due to run."""
        now = datetime.datetime.now()
        
        for job in self.scheduled_campaigns[:]:
            if now >= job['run_time']:
                if self.running or self.followup_running:
                    print(f"[SCHEDULER] A campaign is already running. Rescheduling '{job['campaign_name']}' for one minute later.")
                    job['run_time'] = now + datetime.timedelta(minutes=1)
                else:
                    print(f"[SCHEDULER] Starting scheduled campaign: {job['campaign_name']}")
                    self.status_var.set(f"Status: Starting scheduled campaign '{job['campaign_name']}'...")
                    self.scheduled_campaigns.remove(job) 
                    
                    self.campaign_thread = threading.Thread(
                        target=self.run_campaign_thread,
                        args=(job['recipients'], job['campaign_name'], job['delay_min'], job['delay_max']),
                        daemon=True
                    )
                    self.campaign_thread.start()
                    break 
        
        self.after(60000, self._check_schedule)

    def _load_all_campaign_logs(self):
        """Loads all campaign logs from the logs directory."""
        logs = {}
        for file in os.listdir(config.LOG_DIR):
            if file.endswith(".json"):
                log_filepath = os.path.join(config.LOG_DIR, file)
                try:
                    with open(log_filepath, 'r', encoding='utf-8') as f:
                        log_data = json.load(f)
                    logs[file] = log_data
                except (json.JSONDecodeError, FileNotFoundError) as e:
                    print(f"Error loading log file {file}: {e}")
        return logs

    def _load_initial_data_async(self):
        """Loads logs and all caches in a separate thread and updates the UI."""
        self.after(0, lambda: self.status_var.set("Status: Loading initial data..."))
        
        self.smtp_cache = self._load_from_file(config.SMTP_FILE)
        self.subjects_cache = self._load_from_file(config.SUBJECTS_FILE)
        self.bodies_cache = self._load_from_file(config.EMAIL_BODIES_FILE)
        self.followup_bodies_cache = self._load_from_file(config.FOLLOWUP_BODIES_FILE)
        self.blacklist_cache = self._load_from_file(config.BLACKLIST_FILE)
        self.notifications_cache = self._load_from_file(config.NOTIFICATIONS_FILE)
        
        unread_count = sum(1 for n in self.notifications_cache if not n.get('seen', False))
        
        logs = self._load_all_campaign_logs()
        
        self.after(0, self._update_initial_ui, logs, unread_count)

    def _update_initial_ui(self, logs, unread_count):
        """Callback to update in-memory logs and refresh UI after async load."""
        self.all_campaign_logs = logs
        self.new_notifications_count.set(unread_count) 
        self.status_var.set("Status: All logs and data loaded and ready.")
        self.show_dashboard_ui()

    def _update_live_ui(self, sent, failed, total, index, campaign_id, campaign_name):
        """Updates UI elements with live campaign progress."""
        if self.progress_bar and self.progress_bar.winfo_exists():
            if total > 0:
                self.progress_bar.set((index + 1) / total)
            self.status_var.set(f"Campaign in progress | Sent: {sent}, Failed: {failed}, Total: {total} | Remaining: {total - (sent + failed)}")
        
        if self.analytics_tree and self.analytics_tree.winfo_exists():
            if campaign_id not in self.all_campaign_logs:
                self.all_campaign_logs[campaign_id] = {
                    'id': campaign_id, 'name': campaign_name, 'total_sent': sent,
                    'total_failed': failed, 'timestamp_start': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            self.all_campaign_logs[campaign_id]['total_sent'] = sent
            self.all_campaign_logs[campaign_id]['total_failed'] = failed
            
            if campaign_id in self.analytics_tree.get_children():
                progress_text = f"{int((sent / total) * 100)}% ({sent}/{total})" if total > 0 else "0%"
                self.analytics_tree.item(campaign_id, values=(
                    campaign_name, sent, failed, 
                    self.all_campaign_logs[campaign_id]['timestamp_start'].split(' ')[0], progress_text
                ))
            else:
                self._update_analytics_table()

        if self.viewing_campaign_details_id == campaign_id:
            if hasattr(self, 'email_tree') and self.email_tree and self.email_tree.winfo_exists():
                try:
                    latest_email_entry = self.all_campaign_logs[campaign_id]['emails'][-1]
                    followup_status = latest_email_entry.get('followup_status', 'Not Sent')
                    row_values = (
                        latest_email_entry.get('recipient'),
                        latest_email_entry.get('smtp_used'),
                        latest_email_entry.get('status'),
                        latest_email_entry.get('reason'),
                        followup_status,
                        latest_email_entry.get('followup_count', 0)
                    )
                    self.email_tree.insert("", 0, values=row_values)
                except (KeyError, IndexError) as e:
                    print(f"Could not update detailed view in real-time: {e}")

    def _update_followup_live_ui(self, checked, sent, failed, total):
        """Updates UI elements with live follow-up progress."""
        if self.followup_progress_bar and self.followup_progress_bar.winfo_exists():
            if total > 0:
                self.followup_progress_bar.set(checked / total)
            self.status_var.set(f"Follow-up in progress | Checked: {checked}/{total}, Sent: {sent}, Failed: {failed}")
            
    def _load_from_file(self, filepath):
        """Helper to load JSON file and handle errors."""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            if filepath == config.BLACKLIST_FILE:
                return {}
            return []

    def load_json(self, filepath, cache_key=None):
        """Loads JSON data, using cache if available."""
        if cache_key and getattr(self, cache_key) is not None:
            return getattr(self, cache_key)
        
        data = self._load_from_file(filepath)
        if cache_key:
            setattr(self, cache_key, data)
        return data

    def save_json(self, filepath, data, cache_key=None):
        """Saves JSON data and updates cache."""
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        if cache_key:
            setattr(self, cache_key, data)
            
    def _convert_plain_text_to_html(self, text_content):
        html_paragraphs = []
        for paragraph in text_content.split('\n\n'):
            if paragraph.strip():
                html_paragraph = paragraph.replace('\n', '<br>')
                html_paragraphs.append(f"<p>{html_paragraph}</p>")
        return "\n".join(html_paragraphs)


    def _get_smtp_account_by_email(self, email):
        if self.smtp_cache is None:
            self.smtp_cache = self.load_json(config.SMTP_FILE, 'smtp_cache')
        for smtp in self.smtp_cache:
            if smtp['email'] == email:
                return smtp
        return None

    def _has_replied_in_session(self, imap_session, email_entry):
        if 'message_id' not in email_entry or not email_entry.get('message_id'):
            return False, None
            
        try:
            sent_date = datetime.datetime.strptime(email_entry['timestamp'], "%Y-%m-%d %H:%M:%S").strftime("%d-%b-%Y")
            search_criteria = f'(SINCE {sent_date} (OR HEADER In-Reply-To "{email_entry["message_id"]}" HEADER References "{email_entry["message_id"]}"))'
            status, messages = imap_session.search(None, search_criteria)
            if status == 'OK' and messages[0]:
                return True, messages[0].split()[-1] # Return True and the UID of the reply
        except Exception as e:
            print(f"IMAP check failed for {email_entry['recipient']}: {e}")
        return False, None

    def _run_follow_up_campaign(self, campaign_id):
        self.followup_running = True
        log_data = None
        for file_name, data in self.all_campaign_logs.items():
            if campaign_id in file_name:
                log_data = data
                break 

        if not log_data:
            self.after(0, lambda: messagebox.showerror("Error", f"Could not find campaign data for ID: {campaign_id}"))
            self.followup_running = False
            return
        
        self.after(0, self.show_follow_up_ui)

        followup_bodies = self.load_json(config.FOLLOWUP_BODIES_FILE, 'followup_bodies_cache')
        if not followup_bodies:
            self.after(0, lambda: messagebox.showerror("Error", "Please add follow-up bodies in the Templates section."))
            self.followup_running = False
            return

        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')
        all_eligible_recipients = []
        for email_entry in log_data.get('emails', []):
            is_blacklisted = email_entry.get('recipient') in self.blacklist_cache
            if email_entry.get('status') == 'sent' and not email_entry.get('flag_no_followup') and email_entry.get('followup_status') != 'Replied' and not is_blacklisted:
                all_eligible_recipients.append(email_entry)

        total_to_check = len(all_eligible_recipients)
        self.active_followup_info = {'checked': 0, 'sent': 0, 'failed': 0, 'total': total_to_check}
        
        recipients_by_smtp = defaultdict(list)
        for entry in all_eligible_recipients:
            recipients_by_smtp[entry['smtp_used']].append(entry)

        log_file = os.path.join(config.LOG_DIR, log_data['id'])
        try:
            for smtp_email, recipients in recipients_by_smtp.items():
                if not self.followup_running: break
                
                smtp_account = self._get_smtp_account_by_email(smtp_email)
                if not smtp_account or not smtp_account.get('imap_server'):
                    print(f"[SKIPPING]: No IMAP server configured for {smtp_email}.")
                    self.active_followup_info['checked'] += len(recipients)
                    continue

                imap = None
                try:
                    imap = imaplib.IMAP4_SSL(smtp_account['imap_server'])
                    imap.login(smtp_account['email'], smtp_account['password'])
                    imap.select("inbox")
                    
                    for recipient_entry in recipients:
                        if not self.followup_running: break
                        
                        self.active_followup_info['checked'] += 1
                        
                        has_replied, _ = self._has_replied_in_session(imap, recipient_entry)
                        if has_replied:
                            recipient_entry['followup_status'] = 'Replied'
                        else:
                            current_followup_count = recipient_entry.get('followup_count', 0)
                            template_index = min(current_followup_count, len(followup_bodies) - 1)
                            followup_body_info = followup_bodies[template_index]
                            
                            subject = f"Re: {recipient_entry['subject']}"
                            body_path = os.path.join(config.BODIES_DIR, followup_body_info['file'])
                            
                            try:
                                with open(body_path, 'r', encoding='utf-8') as f:
                                    html_body = f.read()
                                
                                reply_to_id = recipient_entry.get('last_followup_message_id') or recipient_entry.get('message_id')
                                
                                new_message_id = self.send_email(smtp_account, recipient_entry['recipient'], subject, html_body, original_message_id=reply_to_id)

                                if new_message_id:
                                    recipient_entry['followup_status'] = 'Sent'
                                    recipient_entry['followup_count'] = current_followup_count + 1
                                    recipient_entry['last_followup_message_id'] = new_message_id
                                    self.active_followup_info['sent'] += 1
                                else:
                                    recipient_entry['followup_status'] = 'Failed'
                                    self.active_followup_info['failed'] += 1
                            except Exception as e:
                                print(f"Error sending follow-up to {recipient_entry['recipient']}: {e}")
                                recipient_entry['followup_status'] = 'Failed'
                                self.active_followup_info['failed'] += 1
                        
                        self.after(0, self._update_followup_live_ui, 
                                   self.active_followup_info['checked'],
                                   self.active_followup_info['sent'],
                                   self.active_followup_info['failed'],
                                   self.active_followup_info['total'])
                        
                        time.sleep(random.uniform(5, 10))
                        
                except Exception as e:
                    print(f"IMAP or SMTP process failed for {smtp_email}: {e}")
                    self.after(0, lambda err_msg=e: messagebox.showerror("Error", f"IMAP or SMTP error for {smtp_email}: {err_msg}"))
                finally:
                    if imap: imap.logout()
                
                self.save_json(log_file, log_data)

        finally:
            self.followup_running = False
            self.after(0, lambda: self.status_var.set("Follow-up process complete. Refreshing data..."))
            self.after(100, self._refresh_campaign_logs)
            self.after(200, self.show_follow_up_ui)
            
    def _refresh_campaign_logs(self):
        """Refreshes the in-memory log cache and updates relevant UI tables."""
        self.status_var.set("Status: Refreshing all campaign data from disk...")
        self.all_campaign_logs = self._load_all_campaign_logs()
        if hasattr(self, 'analytics_tree') and self.analytics_tree and self.analytics_tree.winfo_exists():
            self._update_analytics_table()
        if hasattr(self, 'followup_campaign_tree') and self.followup_campaign_tree and self.followup_campaign_tree.winfo_exists():
            self._update_followup_campaign_list()
        self.status_var.set("Status: Data refreshed.")


    # ------------------------- MAIN UI SETUP & OTHER METHODS ------------------------- #
    def setup_ui(self):
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="#34495e", height=60)
        self.navigation_frame.grid(row=0, column=0, sticky="ew")
        self.navigation_frame.grid_columnconfigure(0, weight=1)

        # --- Generalized Title ---
        title_label = ctk.CTkLabel(self.navigation_frame, text="Bulk Email Sender", font=ctk.CTkFont(size=20, weight="bold"))
        title_label.grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        button_frame = ctk.CTkFrame(self.navigation_frame, fg_color="transparent")
        button_frame.grid(row=0, column=1, padx=20, pady=5, sticky="e")
        
        self.buttons = {}
        button_info = [
            ("Dashboard", self.show_dashboard_ui),
            ("Campaign", self.show_campaign_ui),
            ("SMTP", self.show_smtp_ui),
            ("Templates", self.show_templates_ui),
            ("Analytics", self.show_analytics_ui),
            ("Follow-up", self.show_follow_up_ui),
            ("Notifications", self.show_notifications_ui)
        ]
        
        for i, (text, command) in enumerate(button_info):
            if text == "Notifications":
                self.notification_button_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
                self.notification_button_frame.grid(row=0, column=i, padx=5)
                
                def update_notification_button_text(count):
                    base_text = "Notifications"
                    return f"• {base_text}" if count > 0 else base_text
                
                self.notification_button = ctk.CTkButton(
                    self.notification_button_frame, 
                    text=update_notification_button_text(self.new_notifications_count.get()), 
                    command=lambda cmd=command: self._run_async(cmd), 
                    fg_color="transparent", hover_color="#2980b9", text_color="white"
                )
                self.notification_button.pack()

                def on_count_change(*args):
                    new_text = update_notification_button_text(self.new_notifications_count.get())
                    self.notification_button.configure(text=new_text)

                self.new_notifications_count.trace_add("write", on_count_change)
            else:
                self.buttons[text] = ctk.CTkButton(button_frame, text=text, command=lambda cmd=command: self._run_async(cmd), fg_color="transparent", hover_color="#2980b9", text_color="white")
                self.buttons[text].grid(row=0, column=i, padx=5)


        self.content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self.status_bar = ctk.CTkLabel(self, textvariable=self.status_var, height=30, corner_radius=0, fg_color="#2a2d2e", font=("Arial", 12))
        self.status_bar.grid(row=2, column=0, sticky="ew")

    def _run_async(self, func):
        """Starts a given function in a new thread to prevent UI from freezing."""
        self.clear_content()
        self.status_var.set("Status: Loading content. Please wait...")
        loading_label = ctk.CTkLabel(self.content_frame, text="Loading...", font=ctk.CTkFont(size=24))
        loading_label.pack(expand=True, fill="both")
        
        self.after(10, lambda: [loading_label.destroy(), func()])

    def clear_content(self):
        self.viewing_campaign_details_id = None
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def strip_html_tags(self, html_text):
        clean = re.compile('<.*?>')
        return re.sub(clean, '', html_text)

    def send_email(self, smtp, to_email, subject, content, original_message_id=None):
        """
        Sends a single email and returns the generated Message-ID on success, or None on failure.
        """
        server = None
        msg_uuid = str(uuid.uuid4())
        domain = smtp['email'].split('@')[1]
        new_message_id = f"<{msg_uuid}@{domain}>"

        try:
            from_email_with_name = f"{smtp.get('name', smtp['email'])} <{smtp['email']}>"
            
            msg = MIMEMultipart("alternative")
            msg['From'] = from_email_with_name
            msg['To'] = to_email
            msg['Subject'] = subject
            msg['Message-ID'] = new_message_id 
            
            if original_message_id:
                msg['In-Reply-To'] = original_message_id
                msg['References'] = original_message_id
            
            if content.strip().startswith('<'):
                plain_text_body = self.strip_html_tags(content)
                html_body = content
            else:
                plain_text_body = content
                html_body = self._convert_plain_text_to_html(content)
                
            part1 = MIMEText(plain_text_body, 'plain')
            part2 = MIMEText(html_body, 'html')
            
            msg.attach(part1)
            msg.attach(part2)
            
            smtp_host = smtp.get('smtp_host', 'smtp.gmail.com')
            smtp_port = smtp.get('smtp_port', 587)
            
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=10)
            server.starttls()
            server.login(smtp['email'], smtp['password'])
            server.sendmail(smtp['email'], to_email, msg.as_string())
            return new_message_id # Return the ID on success
        except Exception as e:
            print(f"ERROR in send_email to {to_email}: {e}")
            return None # Return None on failure
        finally:
            if server:
                try:
                    server.quit()
                except Exception:
                    pass

    def run_campaign_thread(self, recipients, campaign_name, delay_min, delay_max, is_resume=False, log_data=None):
        """
        Main thread function to run or resume an email campaign.
        """
        self.running = True
        
        if is_resume and log_data:
            campaign_file_name = log_data['id']
            recipients_to_send = [email_entry['recipient'] for email_entry in log_data['emails'] if email_entry['status'] != 'sent']
            sent = log_data['total_sent']
            failed = log_data['total_failed']
            total_recipients = len(log_data.get('emails', []))
            
            self.active_campaign_info = {
                'name': campaign_name, 'sent': sent, 'failed': failed,
                'total': total_recipients, 'id': campaign_file_name
            }
        else:
            recipients_to_send = recipients
            sent = 0
            failed = 0
            total_recipients = len(recipients)
            campaign_uuid = str(uuid.uuid4())
            campaign_file_name = f"{campaign_name.replace(' ', '_')}-{campaign_uuid}.json"
            log_data = {
                "id": campaign_file_name, "name": campaign_name,
                "timestamp_start": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "emails": []
            }
            
            self.active_campaign_info = {
                'name': campaign_name, 'sent': 0, 'failed': 0,
                'total': total_recipients, 'id': campaign_file_name
            }
        
        self.after(0, self.show_campaign_ui)
        self.after(0, lambda: self.all_campaign_logs.update({campaign_file_name: log_data}))

        smtps = self.load_json(config.SMTP_FILE, 'smtp_cache')
        subjects = self.load_json(config.SUBJECTS_FILE, 'subjects_cache')
        bodies = self.load_json(config.EMAIL_BODIES_FILE, 'bodies_cache')
        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')

        if not smtps or not subjects or not bodies:
            self.running = False
            self.after(0, lambda: messagebox.showerror("Error", "Please configure SMTP accounts, subjects, and email bodies first."))
            return

        random.shuffle(recipients_to_send)

        try:
            for idx, recipient in enumerate(recipients_to_send):
                if not self.running: break
                
                if recipient in self.blacklist_cache:
                    print(f"Skipping blacklisted recipient: {recipient}")
                    if not is_resume:
                        log_data["emails"].append({
                            "recipient": recipient, "status": "skipped", "reason": "Blacklisted",
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                    continue 

                smtp = smtps[idx % len(smtps)]
                subject = random.choice(subjects)
                body_info = random.choice(bodies)
                
                email_status = "failed"
                reason = "Unknown error"
                message_id = None
                
                try:
                    body_filepath = os.path.join(config.BODIES_DIR, body_info['file'])
                    with open(body_filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    message_id = self.send_email(smtp, recipient, subject, content)
                    
                    if message_id:
                        email_status = "sent"
                        reason = "N/A"
                    else:
                        email_status = "failed"
                        reason = "SMTP error"
                except FileNotFoundError:
                    reason = f"Body file not found: {body_info['file']}"
                except Exception as e:
                    reason = f"An error occurred: {e}"

                email_entry_found = False
                for entry in log_data['emails']:
                    if entry['recipient'] == recipient:
                        entry.update({
                            "status": email_status, "reason": reason,
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "message_id": message_id, "followup_status": "Not Sent",
                            "followup_count": 0, "flag_no_followup": False
                        })
                        email_entry_found = True
                        break
                if not email_entry_found:
                    log_data["emails"].append({
                        "recipient": recipient, "smtp_used": smtp['email'], "subject": subject,
                        "body_template_name": body_info['name'], "status": email_status,
                        "reason": reason, "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "message_id": message_id, "followup_status": "Not Sent",
                        "followup_count": 0, "flag_no_followup": False
                    })
                    
                if email_status == "sent":
                    sent += 1
                else:
                    failed += 1
                
                log_data['total_sent'] = sent
                log_data['total_failed'] = failed
                self.active_campaign_info.update({'sent': sent, 'failed': failed})
                
                self.after(0, lambda s=sent, f=failed, t=total_recipients, i=idx, c_id=campaign_file_name, c_name=campaign_name: self._update_live_ui(s, f, t, i, c_id, c_name))
                log_file = os.path.join(config.LOG_DIR, campaign_file_name)
                self.save_json(log_file, log_data)
                
                delay = random.uniform(delay_min, delay_max)
                time.sleep(delay)
        
        finally:
            self.running = False
            log_data["timestamp_end"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            log_file = os.path.join(config.LOG_DIR, campaign_file_name)
            try:
                self.save_json(log_file, log_data)
                self.after(0, self._refresh_campaign_logs)
            except IOError as e:
                self.after(0, lambda: messagebox.showerror("Log Save Error", f"Could not save campaign log: {e}"))

            self.after(0, lambda: self.status_var.set(f"Campaign finished! Sent: {sent}, Failed: {failed}"))
            self.after(0, lambda: self.show_dashboard_ui())
            
    def _check_for_resumable_campaign(self):
        """Finds the last campaign that was stopped before completion."""
        resumable_campaigns = []
        logs = self._load_all_campaign_logs()
        for file_name, log_data in logs.items():
            if 'timestamp_end' not in log_data and log_data.get('total_sent', 0) + log_data.get('total_failed', 0) < len(log_data.get('emails', [])):
                resumable_campaigns.append((file_name, log_data))
        
        if resumable_campaigns:
            resumable_campaigns.sort(key=lambda x: datetime.datetime.strptime(x[1]['timestamp_start'], "%Y-%m-%d %H:%M:%S"), reverse=True)
            return resumable_campaigns[0]
        
        return None, None

    def show_dashboard_ui(self):
        self.clear_content()
        self.status_var.set("Status: Ready")
        # --- Generalized Text ---
        ctk.CTkLabel(self.content_frame, text="Welcome to the App", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=10)
        ctk.CTkLabel(self.content_frame, text="Professional Campaign Management", font=ctk.CTkFont(size=14, weight="normal")).pack()
        
        self.stats_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.stats_frame.pack(pady=20, fill="x", padx=20)
        self.stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="group1")

        num_smtps = len(self.load_json(config.SMTP_FILE, 'smtp_cache'))
        num_templates = len(self.load_json(config.EMAIL_BODIES_FILE, 'bodies_cache'))
        num_subjects = len(self.load_json(config.SUBJECTS_FILE, 'subjects_cache'))
        num_active_campaigns = 1 if self.running or self.followup_running else 0
        num_scheduled_campaigns = len(self.scheduled_campaigns)

        self.create_stat_card(self.stats_frame, "SMTP Accounts", num_smtps, "#2980b9", 0)
        self.create_stat_card(self.stats_frame, "Email Templates", num_templates, "#27ae60", 1)
        self.create_stat_card(self.stats_frame, "Subject Lines", num_subjects, "#f39c12", 2)
        self.create_stat_card(self.stats_frame, "Active/Scheduled", f"{num_active_campaigns}/{num_scheduled_campaigns}", "#8e44ad", 3)

        ctk.CTkLabel(self.content_frame, text="Quick Actions", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(20, 10))
        quick_actions_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        quick_actions_frame.pack(pady=10, fill="x", padx=20)
        quick_actions_frame.grid_columnconfigure((0, 1), weight=1, uniform="group2")
        quick_actions_frame.grid_rowconfigure((0, 1), weight=1)

        ctk.CTkButton(quick_actions_frame, text="Start New Campaign", command=lambda: self._run_async(self.show_campaign_ui), fg_color="#2ecc71", hover_color="#27ae60").grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        ctk.CTkButton(quick_actions_frame, text="Manage SMTP", command=lambda: self._run_async(self.show_smtp_ui), fg_color="#2980b9", hover_color="#3498db").grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        ctk.CTkButton(quick_actions_frame, text="Edit Templates", command=lambda: self._run_async(self.show_templates_ui), fg_color="#f39c12", hover_color="#e67e22").grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        ctk.CTkButton(quick_actions_frame, text="View Analytics", command=lambda: self._run_async(self.show_analytics_ui), fg_color="#8e44ad", hover_color="#9b59b6").grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
        
        ctk.CTkLabel(self.content_frame, text="Getting Started Checklist", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(20, 10))
        checklist_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        checklist_frame.pack(fill="x", padx=20, pady=5)
        
        has_smtps = "✔" if len(self.load_json(config.SMTP_FILE, 'smtp_cache')) > 0 else "✘"
        has_subjects = "✔" if len(self.load_json(config.SUBJECTS_FILE, 'subjects_cache')) > 0 else "✘"
        has_bodies = "✔" if len(self.load_json(config.EMAIL_BODIES_FILE, 'bodies_cache')) > 0 else "✘"
        
        ctk.CTkLabel(checklist_frame, text=f"{has_smtps} Add at least one SMTP account", font=("Arial", 14)).pack(anchor="w")
        ctk.CTkLabel(checklist_frame, text=f"{has_subjects} Add at least one subject line", font=("Arial", 14)).pack(anchor="w")
        ctk.CTkLabel(checklist_frame, text=f"{has_bodies} Add at least one email template", font=("Arial", 14)).pack(anchor="w")

        detail_button_frame = ctk.CTkFrame(self.content_frame)
        detail_button_frame.pack(side="bottom", fill="x", padx=20, pady=(10,0))
        
        ctk.CTkButton(detail_button_frame, text="See All Campaign Data", 
                      command=lambda: self._run_async(self.show_master_log_viewer_ui)).pack(side="right")

    def show_master_log_viewer_ui(self):
        self.clear_content()
        self.status_var.set("Status: Viewing all campaign data.")
        
        self.content_frame.grid_columnconfigure(0, weight=1, minsize=300)
        self.content_frame.grid_columnconfigure(1, weight=3)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=0)

        left_frame = ctk.CTkFrame(self.content_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        ctk.CTkLabel(left_frame, text="Campaigns", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)

        right_frame = ctk.CTkFrame(self.content_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        
        placeholder_label = ctk.CTkLabel(right_frame, text="Select a campaign from the list to see its details.", font=ctk.CTkFont(size=14))
        placeholder_label.pack(expand=True, padx=20, pady=20)
        
        campaign_list_tree = ttk.Treeview(left_frame, columns=("Name", "Date"), show='headings')
        campaign_list_tree.heading("Name", text="Campaign Name")
        campaign_list_tree.heading("Date", text="Date")
        campaign_list_tree.column("Date", width=100, anchor="center")
        campaign_list_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        sorted_campaigns = sorted(self.all_campaign_logs.values(), key=lambda x: x.get('timestamp_start', '0'), reverse=True)
        for campaign in sorted_campaigns:
            campaign_id = campaign.get('id')
            name = campaign.get('name', 'N/A')
            date_str = campaign.get('timestamp_start', 'N/A').split(' ')[0]
            campaign_list_tree.insert("", "end", iid=campaign_id, values=(name, date_str))

        def on_campaign_select(event):
            for widget in right_frame.winfo_children():
                widget.destroy()

            selected_items = campaign_list_tree.selection()
            if not selected_items:
                placeholder_label = ctk.CTkLabel(right_frame, text="Select a campaign from the list to see its details.", font=ctk.CTkFont(size=14))
                placeholder_label.pack(expand=True, padx=20, pady=20)
                return

            selected_id_from_tree = selected_items[0]
            log_data = self.all_campaign_logs.get(selected_id_from_tree)

            if not log_data:
                return

            scrollable_details = ctk.CTkScrollableFrame(right_frame)
            scrollable_details.pack(fill="both", expand=True, padx=10, pady=10)

            ctk.CTkLabel(scrollable_details, text="Campaign Summary", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 10))
            
            meta_frame = ctk.CTkFrame(scrollable_details, fg_color="transparent")
            meta_frame.pack(fill="x", anchor="w")
            
            meta_data = {
                "Name:": log_data.get('name', 'N/A'),
                "ID:": log_data.get('id', 'N/A'),
                "Started:": log_data.get('timestamp_start', 'N/A'),
                "Finished:": log_data.get('timestamp_end', 'Not Finished'),
                "Total Sent:": log_data.get('total_sent', 0),
                "Total Failed:": log_data.get('total_failed', 0)
            }
            
            for i, (key, value) in enumerate(meta_data.items()):
                ctk.CTkLabel(meta_frame, text=key, font=ctk.CTkFont(weight="bold")).grid(row=i, column=0, sticky="w", padx=5, pady=2)
                ctk.CTkLabel(meta_frame, text=value, wraplength=500).grid(row=i, column=1, sticky="w", padx=5, pady=2)

            ctk.CTkLabel(scrollable_details, text="Email Log", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(20, 10))
            email_log_frame = ctk.CTkFrame(scrollable_details)
            email_log_frame.pack(fill="both", expand=True)

            all_columns = []
            if log_data.get('emails'):
                column_set = set()
                for entry in log_data['emails']:
                    column_set.update(entry.keys())
                
                preferred_order = ['recipient', 'status', 'reason', 'timestamp', 'subject', 'smtp_used', 'message_id', 'followup_status', 'followup_count', 'body_template_name', 'flag_no_followup']
                all_columns = [col for col in preferred_order if col in column_set]
                other_columns = sorted([col for col in column_set if col not in preferred_order])
                all_columns.extend(other_columns)

            if all_columns:
                email_tree = ttk.Treeview(email_log_frame, columns=all_columns, show='headings')
                for col in all_columns:
                    email_tree.heading(col, text=col.replace('_', ' ').title())
                    email_tree.column(col, width=120) 
                
                email_tree.pack(side="left", fill="both", expand=True)
                
                vsb = ctk.CTkScrollbar(email_log_frame, command=email_tree.yview)
                vsb.pack(side="right", fill="y")
                email_tree.configure(yscrollcommand=vsb.set)
                
                hsb = ctk.CTkScrollbar(email_log_frame, command=email_tree.xview, orientation="horizontal")
                hsb.pack(side="bottom", fill="x")
                email_tree.configure(xscrollcommand=hsb.set)
                
                for entry in log_data.get('emails', []):
                    row_values = [entry.get(col, '') for col in all_columns]
                    email_tree.insert("", "end", values=row_values)
            else:
                ctk.CTkLabel(email_log_frame, text="No email entries found for this campaign.").pack(pady=10)

        campaign_list_tree.bind("<<TreeviewSelect>>", on_campaign_select)

    def create_stat_card(self, parent, text, value, color, column):
        card = ctk.CTkFrame(parent, fg_color=color, corner_radius=10)
        card.grid(row=0, column=column, padx=10, pady=10, sticky="nsew")
        ctk.CTkLabel(card, text=str(value), font=ctk.CTkFont(size=36, weight="bold")).pack(pady=(10, 0))
        ctk.CTkLabel(card, text=text, font=ctk.CTkFont(size=16, weight="normal")).pack(pady=(0, 10))

    def show_smtp_ui(self):
        self.clear_content()
        self.status_var.set("Status: Ready to manage SMTP accounts")
        ctk.CTkLabel(self.content_frame, text="Manage SMTPs", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        input_frame = ctk.CTkFrame(self.content_frame)
        input_frame.pack(padx=20, pady=10)

        ctk.CTkLabel(input_frame, text="Display Name:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        # --- Generalized Placeholder ---
        name_var = tk.StringVar(value="[Your Name]")
        name_entry = ctk.CTkEntry(input_frame, textvariable=name_var, width=250)
        name_entry.grid(row=0, column=1, padx=10, pady=5)

        ctk.CTkLabel(input_frame, text="Email:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        email_var = tk.StringVar()
        email_entry = ctk.CTkEntry(input_frame, textvariable=email_var, width=250)
        email_entry.grid(row=1, column=1, padx=10, pady=5)

        ctk.CTkLabel(input_frame, text="App Password:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        password_var = tk.StringVar()
        password_entry = ctk.CTkEntry(input_frame, textvariable=password_var, show="*", width=250)
        password_entry.grid(row=2, column=1, padx=10, pady=5)
        
        smtp_var = tk.StringVar(value="smtp.gmail.com")
        smtp_port_var = tk.StringVar(value="587")
        
        ctk.CTkLabel(input_frame, text="IMAP Server:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        imap_var = tk.StringVar(value="imap.gmail.com")
        imap_entry = ctk.CTkEntry(input_frame, textvariable=imap_var, width=250)
        imap_entry.grid(row=3, column=1, padx=10, pady=5)

        def add_smtp():
            name = name_var.get().strip()
            email = email_var.get().strip()
            password = password_var.get().strip()
            imap_server = imap_var.get().strip()
            smtp_host = smtp_var.get().strip()
            smtp_port = smtp_port_var.get().strip()
            if not name or not email or not password or not imap_server:
                messagebox.showerror("Error", "Please fill all fields.")
                return
            if "@" not in email or "." not in email:
                messagebox.showerror("Error", "Please enter a valid email address.")
                return
            new_entry = {"name": name, "email": email, "password": password, "imap_server": imap_server, "smtp_host": smtp_host, "smtp_port": int(smtp_port)}
            smtps = self.load_json(config.SMTP_FILE, 'smtp_cache')
            smtps.append(new_entry)
            self.save_json(config.SMTP_FILE, smtps, 'smtp_cache')
            self.status_var.set("Status: SMTP account added successfully.")
            self.show_smtp_ui()

        ctk.CTkButton(input_frame, text="Add SMTP", command=add_smtp).grid(row=4, column=1, padx=10, pady=10, sticky="w")

        tree_frame = ctk.CTkFrame(self.content_frame)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2a2d2e", foreground="#dce4ee", fieldbackground="#2a2d2e", borderwidth=0)
        style.configure("Treeview.Heading", background="#343638", foreground="#dce4ee", font=("Arial", 12, "bold"))
        style.map('Treeview', background=[('selected', '#565b5e')])
        
        tree = ttk.Treeview(tree_frame, columns=("Name", "Email", "Password", "IMAP"), show='headings', selectmode="extended")
        tree.heading("Name", text="Display Name")
        tree.heading("Email", text="Email")
        tree.heading("Password", text="Password")
        tree.heading("IMAP", text="IMAP Server")
        tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        scrollbar = ctk.CTkScrollbar(tree_frame, command=tree.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        tree.configure(yscrollcommand=scrollbar.set)
        
        def populate_table():
            for row in tree.get_children():
                tree.delete(row)
            for i, smtp in enumerate(self.load_json(config.SMTP_FILE, 'smtp_cache')):
                tree.insert("", "end", iid=i, values=(smtp.get('name', ''), smtp['email'], smtp['password'], smtp.get('imap_server', '')))
            self.status_var.set(f"Status: {len(self.load_json(config.SMTP_FILE, 'smtp_cache'))} SMTP accounts loaded.")
        
        def test_selected_connection():
            selected = tree.selection()
            if len(selected) != 1:
                messagebox.showerror("Error", "Select a single entry to test.")
                return

            index = int(selected[0])
            smtps = self.load_json(config.SMTP_FILE, 'smtp_cache')
            if index >= len(smtps):
                messagebox.showerror("Error", "Invalid selection.")
                return
            
            smtp_account = smtps[index]
            self.status_var.set(f"Status: Testing connection for {smtp_account['email']}...")

            threading.Thread(target=_test_connection_thread, args=(smtp_account,), daemon=True).start()

        def _test_connection_thread(smtp_account):
            server = None
            try:
                server = smtplib.SMTP(smtp_account.get('smtp_host', 'smtp.gmail.com'), smtp_account.get('smtp_port', 587), timeout=10)
                server.starttls()
                server.login(smtp_account['email'], smtp_account['password'])
                self.after(0, lambda: self.status_var.set(f"Status: Connection to {smtp_account['email']} successful!"))
            except Exception as e:
                self.after(0, lambda err_msg=e: self.status_var.set(f"Status: Connection to {smtp_account['email']} failed. Error: {err_msg}"))
                self.after(0, lambda err_msg=e: messagebox.showerror("Connection Failed", f"Could not connect to {smtp_account['email']}. Please check your credentials and app password.\nError: {err_msg}"))
            finally:
                if server:
                    server.quit()
        
        def delete_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showerror("Error", "Select an entry to delete.")
                return
            if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected SMTP accounts?"):
                smtps = self.load_json(config.SMTP_FILE, 'smtp_cache')
                indices_to_delete = sorted([int(s) for s in selected], reverse=True)
                
                new_smtps = [smtp for i, smtp in enumerate(smtps) if i not in indices_to_delete]

                self.save_json(config.SMTP_FILE, new_smtps, 'smtp_cache')
                populate_table()
                self.status_var.set("Status: Selected SMTP accounts deleted.")
        
        def edit_selected():
            selected = tree.selection()
            if len(selected) != 1:
                messagebox.showerror("Error", "Select a single entry to edit.")
                return
            
            index = int(selected[0])
            current_smtps = self.load_json(config.SMTP_FILE, 'smtp_cache')
            if index >= len(current_smtps):
                messagebox.showerror("Error", "Invalid selection.")
                return

            edit_win = ctk.CTkToplevel(self)
            edit_win.title("Edit SMTP")
            edit_win.geometry("350x300")
            
            ctk.CTkLabel(edit_win, text="Display Name:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
            n_var = tk.StringVar(value=current_smtps[index].get('name', ''))
            ctk.CTkEntry(edit_win, textvariable=n_var, width=200).grid(row=0, column=1, padx=10, pady=10)

            ctk.CTkLabel(edit_win, text="Email:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
            e_var = tk.StringVar(value=current_smtps[index]['email'])
            ctk.CTkEntry(edit_win, textvariable=e_var, width=200).grid(row=1, column=1, padx=10, pady=10)
            
            ctk.CTkLabel(edit_win, text="App Password:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
            p_var = tk.StringVar(value=current_smtps[index]['password'])
            ctk.CTkEntry(edit_win, textvariable=p_var, show="*", width=200).grid(row=2, column=1, padx=10, pady=10)
            
            ctk.CTkLabel(edit_win, text="IMAP Server:").grid(row=3, column=0, padx=10, pady=10, sticky="e")
            i_var = tk.StringVar(value=current_smtps[index].get('imap_server', ''))
            ctk.CTkEntry(edit_win, textvariable=i_var, width=200).grid(row=3, column=1, padx=10, pady=10)

            def save_changes():
                name = n_var.get().strip()
                email = e_var.get().strip()
                password = p_var.get().strip()
                imap_server = i_var.get().strip()

                if not name or not email or not password or not imap_server:
                    messagebox.showerror("Error", "All fields must be filled.")
                    return
                if "@" not in email or "." not in email:
                    messagebox.showerror("Error", "Please enter a valid email address.")
                    return

                smtps = self.load_json(config.SMTP_FILE, 'smtp_cache')
                old_entry = smtps[index]
                smtps[index] = {"name": name, "email": email, "password": password, "imap_server": imap_server, 
                                "smtp_host": old_entry.get("smtp_host", "smtp.gmail.com"), 
                                "smtp_port": old_entry.get("smtp_port", 587)}
                self.save_json(config.SMTP_FILE, smtps, 'smtp_cache')
                edit_win.destroy()
                populate_table()
                self.status_var.set("Status: SMTP account updated successfully.")

            ctk.CTkButton(edit_win, text="Save Changes", command=save_changes).grid(row=4, column=1, pady=10)

        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="Delete Selected", command=delete_selected, fg_color="#e74c3c", hover_color="#c0392b").pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Edit Selected", command=edit_selected, fg_color="#f39c12", hover_color="#e67e22").pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Test Selected Connection", command=test_selected_connection, fg_color="#3498db", hover_color="#2980b9").pack(side="left", padx=5)
        
        populate_table()

    def show_templates_ui(self):
        self.clear_content()
        self.status_var.set("Status: Ready to manage email templates and subjects.")
        ctk.CTkLabel(self.content_frame, text="Manage Templates", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        scrollable_content_frame = ctk.CTkScrollableFrame(self.content_frame)
        scrollable_content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Subjects Section
        subj_frame = ctk.CTkFrame(scrollable_content_frame)
        subj_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(subj_frame, text="Subjects:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        subj_input_frame = ctk.CTkFrame(subj_frame, fg_color="transparent")
        subj_input_frame.pack(fill="x", padx=10)
        subj_var = tk.StringVar()
        ctk.CTkEntry(subj_input_frame, textvariable=subj_var, placeholder_text="New Subject").pack(side="left", fill="x", expand=True)
        ctk.CTkButton(subj_input_frame, text="Add", command=lambda: self.add_subject(subj_var.get())).pack(side="left", padx=(10, 0))
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2a2d2e", foreground="#dce4ee", fieldbackground="#2a2d2e", borderwidth=0)
        style.configure("Treeview.Heading", background="#343638", foreground="#dce4ee", font=("Arial", 12, "bold"))
        style.map('Treeview', background=[('selected', '#565b5e')])
        
        self.subj_tree = ttk.Treeview(subj_frame, columns=("Subject",), show='headings', height=5, selectmode="extended")
        self.subj_tree.heading("Subject", text="Subject")
        self.subj_tree.pack(fill="x", padx=10, pady=5)
        ctk.CTkButton(subj_frame, text="Delete Selected", command=self.delete_subject, fg_color="#e74c3c", hover_color="#c0392b").pack(pady=(5, 10))

        # Email Bodies Section
        body_frame = ctk.CTkFrame(scrollable_content_frame)
        body_frame.pack(fill="both", expand=True, padx=20, pady=10)
        ctk.CTkLabel(body_frame, text="Email Bodies:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        body_button_frame = ctk.CTkFrame(body_frame, fg_color="transparent")
        body_button_frame.pack(fill="x", padx=10)
        ctk.CTkButton(body_button_frame, text="Add New Body", command=lambda: self.add_body_wizard(is_followup=False)).pack(side="left", padx=5)
        
        self.body_tree = ttk.Treeview(body_frame, columns=("Name", "Type"), show='headings', selectmode="extended")
        self.body_tree.heading("Name", text="Name")
        self.body_tree.heading("Type", text="Type")
        self.body_tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        button_body_tree_frame = ctk.CTkFrame(body_frame, fg_color="transparent")
        button_body_tree_frame.pack(pady=5)
        ctk.CTkButton(button_body_tree_frame, text="Delete Selected Body", command=lambda: self.delete_body(is_followup=False), fg_color="#e74c3c", hover_color="#c0392b").pack(side="left", padx=5)
        ctk.CTkButton(button_body_tree_frame, text="Edit Selected Body", command=lambda: self._show_edit_window(is_followup=False), fg_color="#f39c12", hover_color="#e67e22").pack(side="left", padx=5)
        
        # Follow-up Bodies Section
        followup_body_frame = ctk.CTkFrame(scrollable_content_frame)
        followup_body_frame.pack(fill="both", expand=True, padx=20, pady=10)
        ctk.CTkLabel(followup_body_frame, text="Follow-up Bodies:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        followup_body_button_frame = ctk.CTkFrame(followup_body_frame, fg_color="transparent")
        followup_body_button_frame.pack(fill="x", padx=10)
        ctk.CTkButton(followup_body_button_frame, text="Add New Follow-up Body", command=lambda: self.add_body_wizard(is_followup=True)).pack(side="left", padx=5)

        self.followup_body_tree = ttk.Treeview(followup_body_frame, columns=("Name", "Type"), show='headings', selectmode="extended")
        self.followup_body_tree.heading("Name", text="Name")
        self.followup_body_tree.heading("Type", text="Type")
        self.followup_body_tree.pack(fill="both", expand=True, padx=10, pady=5)

        button_followup_body_tree_frame = ctk.CTkFrame(followup_body_frame, fg_color="transparent")
        button_followup_body_tree_frame.pack(pady=5)
        ctk.CTkButton(button_followup_body_tree_frame, text="Delete Selected Follow-up Body", command=lambda: self.delete_body(is_followup=True), fg_color="#e74c3c", hover_color="#c0392b").pack(side="left", padx=5)
        ctk.CTkButton(button_followup_body_tree_frame, text="Edit Selected Follow-up Body", command=lambda: self._show_edit_window(is_followup=True), fg_color="#f39c12", hover_color="#e67e22").pack(side="left", padx=5)

        self.populate_templates()
        
    def add_body_wizard(self, is_followup=False):
        """Shows a wizard to choose how to add a new email body."""
        win = ctk.CTkToplevel(self)
        win.title("Add New Email Body")
        win.geometry("400x200")
        
        ctk.CTkLabel(win, text="How would you like to add a new body?", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)
        
        button_frame = ctk.CTkFrame(win, fg_color="transparent")
        button_frame.pack(pady=10)
        
        ctk.CTkButton(button_frame, text="Write from Scratch", command=lambda: [win.destroy(), self.add_body_write(is_followup)]).pack(pady=5)
        ctk.CTkButton(button_frame, text="Upload File", command=lambda: [win.destroy(), self.add_body_upload(is_followup)]).pack(pady=5)
        
    def add_subject(self, subject):
        if not subject.strip():
            messagebox.showerror("Error", "Subject cannot be empty.")
            return
        subjects = self.load_json(config.SUBJECTS_FILE, 'subjects_cache')
        subjects.append(subject.strip())
        self.save_json(config.SUBJECTS_FILE, subjects, 'subjects_cache')
        self.populate_templates()
        self.status_var.set("Status: Subject line added successfully.")

    def delete_subject(self):
        selected = self.subj_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Select a subject to delete.")
            return
        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected subject lines?"):
            subjects = self.load_json(config.SUBJECTS_FILE, 'subjects_cache')
            subjects_to_delete = [self.subj_tree.item(item)['values'][0] for item in selected]
            
            new_subjects = [subj for subj in subjects if subj not in subjects_to_delete]
            
            self.save_json(config.SUBJECTS_FILE, new_subjects, 'subjects_cache')
            self.populate_templates()
            self.status_var.set("Status: Selected subject lines deleted.")

    def add_body_write(self, is_followup=False):
        win = ctk.CTkToplevel(self)
        win.title("Add Email Body")
        win.geometry("800x600")
        
        ctk.CTkLabel(win, text="Body Name:").pack(pady=(10, 5))
        name_var = tk.StringVar()
        ctk.CTkEntry(win, textvariable=name_var).pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(win, text="Body Content (Plain text or HTML):").pack(pady=(10, 5))
        text_frame = ctk.CTkFrame(win)
        text_frame.pack(padx=20, pady=5, fill="both", expand=True)
        text_editor = ctk.CTkTextbox(text_frame)
        text_editor.pack(side="left", fill="both", expand=True)

        preview_frame = ctk.CTkFrame(win, fg_color="gray20")
        preview_frame.pack(padx=20, pady=5, fill="both", expand=True)
        ctk.CTkLabel(preview_frame, text="Live Preview:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        preview_label = ctk.CTkLabel(preview_frame, text="Start typing to see a preview...", wraplength=750, justify="left", text_color="white", anchor="nw")
        preview_label.pack(fill="both", expand=True, padx=5, pady=5)
        
        def update_preview(event=None):
            content = text_editor.get("1.0", "end-1c")
            if "<" in content and ">" in content:
                preview_text = self.strip_html_tags(content)
            else:
                preview_text = content
            
            preview_label.configure(text=preview_text)

        text_editor.bind("<KeyRelease>", update_preview)

        def save():
            if not name_var.get().strip():
                messagebox.showerror("Error", "Body Name cannot be empty.")
                return
            
            filename = f"{uuid.uuid4()}.html"
            content = text_editor.get("1.0", "end-1c")
            
            if "<" in content and ">" in content:
                file_type = "html"
            else:
                file_type = "text"

            try:
                filepath = os.path.join(config.BODIES_DIR, filename)
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content)
            except IOError as e:
                messagebox.showerror("File Error", f"Could not save body file: {e}")
                return
            
            if is_followup:
                bodies = self.load_json(config.FOLLOWUP_BODIES_FILE, 'followup_bodies_cache')
                file_to_save = config.FOLLOWUP_BODIES_FILE
                cache_key = 'followup_bodies_cache'
            else:
                bodies = self.load_json(config.EMAIL_BODIES_FILE, 'bodies_cache')
                file_to_save = config.EMAIL_BODIES_FILE
                cache_key = 'bodies_cache'
            
            bodies.append({"name": name_var.get(), "file": filename, "type": file_type})
            
            self.save_json(file_to_save, bodies, cache_key)
            self.status_var.set(f"Status: Email body added successfully.")
            
            win.destroy()
            self.populate_templates()

        ctk.CTkButton(win, text="Save", command=save).pack(pady=20)

    def add_body_upload(self, is_followup=False):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("HTML files", "*.html")])
        if not file_path:
            return

        name = os.path.basename(file_path)
        ext = os.path.splitext(name)[1].lower()
        if ext not in [".txt", ".html"]:
            messagebox.showerror("Error", "Unsupported file type.")
            return
        
        new_filename = f"{uuid.uuid4()}.html"
        try:
            with open(file_path, 'r', encoding='utf-8') as source_file:
                content = source_file.read()
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='latin-1') as source_file:
                    content = source_file.read()
            except Exception as e:
                messagebox.showerror("File Read Error", f"Could not read file: {e}")
                return
        except FileNotFoundError:
            messagebox.showerror("Error", "File not found.")
            return
        
        if "<" in content and ">" in content:
            file_type = "html"
        else:
            file_type = "text"

        try:
            filepath = os.path.join(config.BODIES_DIR, new_filename)
            with open(filepath, 'w', encoding='utf-8') as dest_file:
                dest_file.write(content)
        except IOError as e:
            messagebox.showerror("File Save Error", f"Could not save body file: {e}")
            return
        
        if is_followup:
            bodies = self.load_json(config.FOLLOWUP_BODIES_FILE, 'followup_bodies_cache')
            file_to_save = config.FOLLOWUP_BODIES_FILE
            cache_key = 'followup_bodies_cache'
        else:
            bodies = self.load_json(config.EMAIL_BODIES_FILE, 'bodies_cache')
            file_to_save = config.EMAIL_BODIES_FILE
            cache_key = 'bodies_cache'
            
        bodies.append({"name": name, "file": new_filename, "type": file_type})
        
        self.save_json(file_to_save, bodies, cache_key)
        self.status_var.set(f"Status: Email body from '{name}' added.")
            
        self.populate_templates()

    def delete_body(self, is_followup=False):
        if is_followup:
            tree = self.followup_body_tree
            file = config.FOLLOWUP_BODIES_FILE
            cache_key = 'followup_bodies_cache'
        else:
            tree = self.body_tree
            file = config.EMAIL_BODIES_FILE
            cache_key = 'bodies_cache'

        selected = tree.selection()
        if not selected:
            messagebox.showerror("Error", "Select an email body to delete.")
            return
        
        if not messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected email bodies?"):
            return

        bodies = self.load_json(file, cache_key)
        
        file_names_to_delete = [bodies[int(item)]['file'] for item in selected if int(item) < len(bodies)]
        
        new_bodies = [body for i, body in enumerate(bodies) if str(i) not in selected]
        
        for file_name in file_names_to_delete:
            filepath = os.path.join(config.BODIES_DIR, file_name)
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except OSError as e:
                    print(f"Error deleting file {filepath}: {e}")
                    messagebox.showerror("File Deletion Error", f"Could not delete file {file_name}: {e}")
        
        self.save_json(file, new_bodies, cache_key)
        self.populate_templates()
        self.status_var.set("Status: Selected email bodies deleted.")
    
    def _show_edit_window(self, is_followup=False):
        if is_followup:
            tree = self.followup_body_tree
            file = config.FOLLOWUP_BODIES_FILE
        else:
            tree = self.body_tree
            file = config.EMAIL_BODIES_FILE

        selected = tree.selection()
        if len(selected) != 1:
            messagebox.showerror("Error", "Select a single entry to edit.")
            return

        index = int(selected[0])
        bodies = self.load_json(file)
        if index >= len(bodies):
            messagebox.showerror("Error", "Invalid selection.")
            return
        
        original_body = bodies[index]
        body_filepath = os.path.join(config.BODIES_DIR, original_body['file'])

        win = ctk.CTkToplevel(self)
        win.title(f"Edit {original_body['name']}")
        win.geometry("800x600")

        ctk.CTkLabel(win, text="Body Name:").pack(pady=(10, 5))
        name_var = tk.StringVar(value=original_body['name'])
        ctk.CTkEntry(win, textvariable=name_var).pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(win, text="Body Content (Plain text or HTML):").pack(pady=(10, 5))
        text_frame = ctk.CTkFrame(win)
        text_frame.pack(padx=20, pady=5, fill="both", expand=True)
        text_editor = ctk.CTkTextbox(text_frame)
        text_editor.pack(side="left", fill="both", expand=True)

        preview_frame = ctk.CTkFrame(win, fg_color="gray20")
        preview_frame.pack(padx=20, pady=5, fill="both", expand=True)
        ctk.CTkLabel(preview_frame, text="Live Preview:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        preview_label = ctk.CTkLabel(preview_frame, text="Start typing to see a preview...", wraplength=750, justify="left", text_color="white", anchor="nw")
        preview_label.pack(fill="both", expand=True, padx=5, pady=5)

        if os.path.exists(body_filepath):
            with open(body_filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            text_editor.insert("1.0", content)
            
            if original_body.get('type') == 'html':
                preview_text = self.strip_html_tags(content)
            else:
                preview_text = content
            
            preview_label.configure(text=preview_text)
        
        def update_preview(event=None):
            content = text_editor.get("1.0", "end-1c")
            
            if original_body.get('type') == 'html':
                preview_text = self.strip_html_tags(content)
            else:
                preview_text = content
            
            preview_label.configure(text=preview_text)

        text_editor.bind("<KeyRelease>", update_preview)

        def save_edit():
            if not name_var.get().strip():
                messagebox.showerror("Error", "Body Name cannot be empty.")
                return

            new_content = text_editor.get("1.0", "end-1c")
            try:
                with open(body_filepath, 'w', encoding='utf-8') as f:
                    f.write(new_content)
            except IOError as e:
                messagebox.showerror("File Error", f"Could not save body file: {e}")
                return
            
            bodies[index]['name'] = name_var.get()
            if "<" in new_content and ">" in new_content:
                bodies[index]['type'] = "html"
            else:
                bodies[index]['type'] = "text"
            
            self.save_json(file, bodies)
            win.destroy()
            self.populate_templates()
            self.status_var.set("Status: Template updated successfully.")
            
        ctk.CTkButton(win, text="Save Changes", command=save_edit).pack(pady=20)
        
    def populate_templates(self):
        for row in self.subj_tree.get_children():
            self.subj_tree.delete(row)
        for i, subj in enumerate(self.load_json(config.SUBJECTS_FILE, 'subjects_cache')):
            self.subj_tree.insert("", "end", iid=i, values=(subj,))
        
        for row in self.body_tree.get_children():
            self.body_tree.delete(row)
        for i, body in enumerate(self.load_json(config.EMAIL_BODIES_FILE, 'bodies_cache')):
            self.body_tree.insert("", "end", iid=i, values=(body['name'], body['type']))
        
        for row in self.followup_body_tree.get_children():
            self.followup_body_tree.delete(row)
        for i, body in enumerate(self.load_json(config.FOLLOWUP_BODIES_FILE, 'followup_bodies_cache')):
            self.followup_body_tree.insert("", "end", iid=i, values=(body['name'], body['type']))
        self.status_var.set("Status: Templates and subjects reloaded.")

    def load_recipients(self, path):
        recipients = []
        if not os.path.exists(path):
            raise FileNotFoundError(f"Recipients file not found at: '{path}'")

        try:
            if path.lower().endswith(".csv"):
                with open(path, newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    email_key = next((k for k in reader.fieldnames if k.lower() == 'email'), None)
                    if not email_key:
                        raise ValueError("CSV file must contain an 'email' column.")
                    
                    for row in reader:
                        email = row.get(email_key, '').strip()
                        if email and '@' in email and '.' in email:
                            recipients.append(email)
            else:
                with open(path, 'r', encoding='utf-8') as f:
                    for line in f:
                        email = line.strip()
                        if email and '@' in email and '.' in email:
                            recipients.append(email)
        except UnicodeDecodeError:
            raise ValueError(f"Could not read file '{path}'. Please ensure it's a valid UTF-8 or plain text file.")
        except Exception as e:
            raise ValueError(f"Error reading recipients file '{path}': {e}")
        
        return recipients

    def show_campaign_ui(self):
        self.clear_content()
        
        if self.running:
            # Display live progress UI if a campaign is running
            self.status_var.set(f"Campaign in progress | Sent: {self.active_campaign_info.get('sent', 0)}, Failed: {self.active_campaign_info.get('failed', 0)}, Total: {self.active_campaign_info.get('total', 0)}")
            
            ctk.CTkLabel(self.content_frame, text="Live Campaign Progress", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
            
            live_update_frame = ctk.CTkFrame(self.content_frame, fg_color="#2a2d2e", corner_radius=10)
            live_update_frame.pack(pady=20, padx=20, fill="x")
            
            self.live_status_label = ctk.CTkLabel(live_update_frame, textvariable=self.status_var, font=("Arial", 14, "bold"))
            self.live_status_label.pack(pady=5)
            
            self.progress_bar = ctk.CTkProgressBar(live_update_frame, width=400)
            if self.active_campaign_info.get('total', 0) > 0:
                self.progress_bar.set((self.active_campaign_info.get('sent', 0) + self.active_campaign_info.get('failed', 0)) / self.active_campaign_info['total'])
            self.progress_bar.pack(pady=10)
            
            def stop_campaign_action():
                self.running = False
                messagebox.showinfo("Campaign Stopped", "The campaign is being stopped. Please wait.")
            
            button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
            button_frame.pack(pady=20)
            
            self.campaign_btn = ctk.CTkButton(button_frame, text="Stop Campaign", command=stop_campaign_action, fg_color="#e74c3c", hover_color="#c0392b")
            self.campaign_btn.pack(side="left", padx=10)

        else:
            resumable_campaign_file, resumable_campaign_data = self._check_for_resumable_campaign()
            
            if resumable_campaign_data:
                # Display resume campaign UI
                self.status_var.set(f"Status: Found a resumable campaign.")
                ctk.CTkLabel(self.content_frame, text="Resume Campaign", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
                
                resume_frame = ctk.CTkFrame(self.content_frame)
                resume_frame.pack(padx=20, pady=10, fill="x")
                
                campaign_name = resumable_campaign_data.get('name', 'Unnamed Campaign')
                total_emails = len(resumable_campaign_data.get('emails', []))
                sent = resumable_campaign_data.get('total_sent', 0)
                failed = resumable_campaign_data.get('total_failed', 0)
                remaining = total_emails - (sent + failed)

                ctk.CTkLabel(resume_frame, text=f"Campaign Name: {campaign_name}", font=("Arial", 14, "bold")).pack(pady=5)
                ctk.CTkLabel(resume_frame, text=f"Last run on: {resumable_campaign_data.get('timestamp_start', 'N/A')}", font=("Arial", 12)).pack(pady=5)
                ctk.CTkLabel(resume_frame, text=f"Progress: {sent} sent, {failed} failed, {remaining} remaining.", font=("Arial", 12)).pack(pady=5)
                
                def resume_action():
                    self.campaign_thread = threading.Thread(
                        target=self.run_campaign_thread,
                        args=(None, campaign_name, 120, 180, True, resumable_campaign_data),
                        daemon=True
                    )
                    self.campaign_thread.start()

                ctk.CTkButton(self.content_frame, text=f"Resume '{campaign_name}'", command=resume_action, fg_color="#2ecc71", hover_color="#27ae60").pack(pady=20)

            else:
                # Display the standard start campaign UI if no campaign is running
                self.status_var.set("Status: Ready to start a new campaign.")
                ctk.CTkLabel(self.content_frame, text="Run Email Campaign", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
                
                main_frame = ctk.CTkFrame(self.content_frame)
                main_frame.pack(padx=20, pady=10, fill="x")
                main_frame.columnconfigure(1, weight=1)

                ctk.CTkLabel(main_frame, text="Campaign Name:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
                campaign_name_entry = ctk.CTkEntry(main_frame, width=300)
                campaign_name_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

                ctk.CTkLabel(main_frame, text="Recipients File (CSV/TXT):").grid(row=1, column=0, padx=10, pady=10, sticky="e")
                recipient_path = tk.StringVar()
                file_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
                file_frame.grid(row=1, column=1, padx=10, pady=10, sticky="w")
                ctk.CTkEntry(file_frame, textvariable=recipient_path, width=250).pack(side="left")
                browse_btn = ctk.CTkButton(file_frame, text="Browse", width=80, command=lambda: recipient_path.set(filedialog.askopenfilename(
                    filetypes=[("Recipient files", "*.csv *.txt"), ("All files", "*.*")]
                )))
                browse_btn.pack(side="left", padx=(10, 0))

                ctk.CTkLabel(main_frame, text="Delay (Min-Max) Seconds:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
                delay_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
                delay_frame.grid(row=2, column=1, padx=10, pady=10, sticky="w")
                delay_min_var = tk.StringVar(value="120")
                delay_max_var = tk.StringVar(value="180")
                ctk.CTkEntry(delay_frame, textvariable=delay_min_var, width=60).pack(side="left")
                ctk.CTkLabel(delay_frame, text=" to ").pack(side="left", padx=5)
                ctk.CTkEntry(delay_frame, textvariable=delay_max_var, width=60).pack(side="left")
                
                schedule_var = tk.BooleanVar()
                schedule_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
                schedule_frame.grid(row=3, column=1, padx=10, pady=10, sticky="w")
                
                schedule_checkbox = ctk.CTkCheckBox(main_frame, text="Schedule for later", variable=schedule_var, command=lambda: toggle_schedule_fields())
                schedule_checkbox.grid(row=3, column=0, padx=10, pady=10, sticky="e")

                date_entry = ctk.CTkEntry(schedule_frame, placeholder_text="YYYY-MM-DD")
                time_entry = ctk.CTkEntry(schedule_frame, placeholder_text="HH:MM (24h)")
                
                def toggle_schedule_fields():
                    if schedule_var.get():
                        date_entry.pack(side="left", padx=(0, 5))
                        time_entry.pack(side="left")
                        campaign_btn.configure(text="Schedule Campaign")
                    else:
                        date_entry.pack_forget()
                        time_entry.pack_forget()
                        campaign_btn.configure(text="Start Campaign")

                def start_or_schedule_campaign():
                    if self.running:
                        messagebox.showinfo("Info", "A campaign is already running.")
                        return

                    campaign_name = campaign_name_entry.get().strip()
                    path = recipient_path.get().strip()
                    
                    if not campaign_name or not path:
                        messagebox.showerror("Missing Input", "Please provide a Campaign Name and Recipients File.")
                        return

                    try:
                        delay_min = float(delay_min_var.get())
                        delay_max = float(delay_max_var.get())
                        recipients = self.load_recipients(path)
                        if not recipients:
                            messagebox.showerror("Error", "No valid recipients found in the file.")
                            return
                    except (ValueError, FileNotFoundError) as e:
                        messagebox.showerror("Input Error", str(e))
                        return

                    if schedule_var.get():
                        date_str = date_entry.get().strip()
                        time_str = time_entry.get().strip()
                        try:
                            run_datetime = datetime.datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
                            if run_datetime < datetime.datetime.now():
                                messagebox.showerror("Error", "Scheduled time cannot be in the past.")
                                return
                            
                            job = {
                                "campaign_name": campaign_name,
                                "recipients": recipients,
                                "delay_min": delay_min,
                                "delay_max": delay_max,
                                "run_time": run_datetime
                            }
                            self.scheduled_campaigns.append(job)
                            messagebox.showinfo("Success", f"Campaign '{campaign_name}' scheduled for {run_datetime.strftime('%Y-%m-%d %I:%M %p')}.")
                            self.show_dashboard_ui()

                        except ValueError:
                            messagebox.showerror("Error", "Invalid date or time format. Use YYYY-MM-DD and HH:MM.")
                            return
                    else:
                        self.campaign_thread = threading.Thread(
                            target=self.run_campaign_thread,
                            args=(recipients, campaign_name, delay_min, delay_max),
                            daemon=True
                        )
                        self.campaign_thread.start()

                button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
                button_frame.pack(pady=20)
                
                campaign_btn = ctk.CTkButton(button_frame, text="Start Campaign", command=start_or_schedule_campaign, fg_color="#2ecc71", hover_color="#27ae60")
                campaign_btn.pack(side="left", padx=10)
    
    def export_campaign_log(self, file_name):
        """Exports a single campaign log to a CSV file."""
        log_data = self.all_campaign_logs.get(file_name)
        if not log_data:
            messagebox.showerror("Error", "Could not find the selected log data in memory.")
            return

        default_filename = f"campaign_log_{log_data.get('name', 'export').replace(' ', '_')}_{log_data.get('id', 'full')}.csv"
        export_filepath = filedialog.asksaveasfilename(defaultextension=".csv",
                                                     initialfile=default_filename,
                                                     filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        
        if not export_filepath:
            return False

        try:
            fieldnames = ['recipient', 'smtp_used', 'subject', 'body_template_name', 'status', 'reason', 'timestamp', 'followup_status', 'followup_count']
            
            with open(export_filepath, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                for email_entry in log_data.get('emails', []):
                    row_to_write = {field: email_entry.get(field, '') for field in fieldnames}
                    writer.writerow(row_to_write)
            
            self.status_var.set(f"Status: Log exported successfully to {os.path.basename(export_filepath)}.")
            messagebox.showinfo("Success", f"Log exported successfully to:\n{export_filepath}")
            return True
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting the log: {e}")
            return False

    def _auto_archive_old_campaigns(self):
        """Finds and archives campaigns older than 3 months."""
        if not messagebox.askyesno("Confirm Auto-Archive", "This will permanently delete campaign data older than 3 months. Do you want to continue?"):
            return

        self.status_var.set("Status: Checking for old campaigns to archive...")
        three_months_ago = datetime.datetime.now() - datetime.timedelta(days=90)
        
        campaigns_to_archive = []
        for file_name, log_data in list(self.all_campaign_logs.items()):
            start_time_str = log_data.get('timestamp_start')
            if start_time_str:
                campaign_date = datetime.datetime.now()
                try:
                    campaign_date = datetime.datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    print(f"Warning: Could not parse date for campaign {file_name}")
                
                if campaign_date < three_months_ago:
                    campaigns_to_archive.append((file_name, log_data))

        if not campaigns_to_archive:
            self.status_var.set("Status: No campaigns older than 3 months were found.")
            messagebox.showinfo("Info", "No campaigns older than 3 months were found.")
            return

        for file_name, log_data in campaigns_to_archive:
            if messagebox.askyesno("Export Before Deleting?", f"Campaign '{log_data['name']}' is older than 3 months. Do you want to export it before deleting?"):
                self.export_campaign_log(file_name)

            log_filepath = os.path.join(config.LOG_DIR, file_name)
            try:
                os.remove(log_filepath)
                if file_name in self.all_campaign_logs:
                    del self.all_campaign_logs[file_name]
                self.status_var.set(f"Status: Archived and deleted log for '{log_data['name']}'.")
            except OSError as e:
                self.status_var.set(f"Error deleting file for '{log_data['name']}': {e}")

        self._update_analytics_table()
        self.status_var.set("Status: Auto-archiving process complete.")

    def delete_selected_campaigns(self):
        """Deletes selected campaign logs from the Analytics table and disk."""
        selected_items = self.analytics_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "Please select one or more campaigns to delete.")
            return

        if not messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete {len(selected_items)} campaign(s) permanently? This action cannot be undone."):
            return

        self.status_var.set("Status: Deleting selected campaigns...")
        
        files_to_delete = list(selected_items)
        for file_name in files_to_delete:
            log_filepath = os.path.join(config.LOG_DIR, file_name)
            try:
                os.remove(log_filepath)
                if file_name in self.all_campaign_logs:
                    del self.all_campaign_logs[file_name]
                self.analytics_tree.delete(file_name)
            except OSError as e:
                self.status_var.set(f"Error deleting file {file_name}: {e}")
                messagebox.showerror("Deletion Error", f"Could not delete file {file_name}: {e}")

        self.status_var.set(f"Status: Deleted {len(files_to_delete)} campaign(s) successfully.")
        self._update_analytics_table()

    def show_analytics_ui(self):
        self.clear_content()
        self.status_var.set("Status: Ready to view campaign analytics.")
        ctk.CTkLabel(self.content_frame, text="Campaign Analytics", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        filter_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        filter_frame.pack(fill="x", padx=20, pady=10)
        
        search_var = tk.StringVar()
        search_entry = ctk.CTkEntry(filter_frame, textvariable=search_var, placeholder_text="Search campaigns...", width=300)
        search_entry.pack(side="left", padx=(0, 10))
        search_entry.bind("<KeyRelease>", lambda event: self._update_analytics_table(search_var.get()))
        
        ctk.CTkButton(filter_frame, text="Search", command=lambda: self._update_analytics_table(search_var.get())).pack(side="left", padx=5)
        ctk.CTkButton(filter_frame, text="Reset", command=lambda: self._update_analytics_table()).pack(side="left", padx=5)
        
        ctk.CTkButton(filter_frame, text="DNC Management", command=self.show_dnc_ui, fg_color="#3498db", hover_color="#2980b9").pack(side="right", padx=5)

        ctk.CTkButton(filter_frame, text="Auto-Archive Old Campaigns", command=self._auto_archive_old_campaigns, fg_color="#e74c3c", hover_color="#c0392b").pack(side="right", padx=10)

        log_frame = ctk.CTkFrame(self.content_frame)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2a2d2e", foreground="#dce4ee", fieldbackground="#2a2d2e", borderwidth=0)
        style.configure("Treeview.Heading", background="#343638", foreground="#dce4ee", font=("Arial", 12, "bold"))
        style.map('Treeview', background=[('selected', '#565b5e')])
        
        self.analytics_tree = ttk.Treeview(log_frame, columns=("Campaign Name", "Sent", "Failed", "Date & Time", "Time Passed"), show='headings', selectmode="extended")
        self.analytics_tree.heading("Campaign Name", text="Campaign Name", anchor="w")
        self.analytics_tree.heading("Sent", text="Sent", anchor="center")
        self.analytics_tree.heading("Failed", text="Failed", anchor="center")
        self.analytics_tree.heading("Date & Time", text="Date & Time", anchor="w")
        self.analytics_tree.heading("Time Passed", text="Time Passed", anchor="w")
        self.analytics_tree.column("Date & Time", width=150)
        self.analytics_tree.column("Time Passed", width=150)
        self.analytics_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        
        scrollbar = ctk.CTkScrollbar(log_frame, command=self.analytics_tree.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        self.analytics_tree.configure(yscrollcommand=scrollbar.set)
        
        self._update_analytics_table()

        def show_detailed_analytics():
            selected_items = self.analytics_tree.selection()
            if not selected_items:
                messagebox.showerror("Error", "Please select a campaign to view analytics.")
                return
            if len(selected_items) > 1:
                messagebox.showerror("Error", "Please select only one campaign at a time.")
                return
            
            file_name = selected_items[0]
            log_data = self.all_campaign_logs.get(file_name)
            
            if not log_data:
                messagebox.showerror("Error", "Could not find the selected log data in memory.")
                return

            self.show_campaign_details(log_data)

        def export_selected_log_action():
            selected_items = self.analytics_tree.selection()
            if not selected_items:
                messagebox.showerror("Error", "Please select a campaign log to export.")
                return
            if len(selected_items) > 1:
                messagebox.showerror("Error", "Please select only one campaign log to export at a time.")
                return
            
            file_name = selected_items[0]
            self.export_campaign_log(file_name)

        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="View Details", command=show_detailed_analytics).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Export Selected Log to CSV", command=export_selected_log_action).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Delete Selected Campaign(s)", command=self.delete_selected_campaigns, fg_color="#e74c3c", hover_color="#c0392b").pack(side="left", padx=5)

    def _update_analytics_table(self, query=""):
        if not self.analytics_tree or not self.analytics_tree.winfo_exists():
            return
            
        for row in self.analytics_tree.get_children():
            self.analytics_tree.delete(row)

        for file, log_data in self.all_campaign_logs.items():
            campaign_name = log_data.get('name', 'Unnamed Campaign')
            if query.lower() in campaign_name.lower():
                total_sent = log_data.get('total_sent', 0)
                total_failed = log_data.get('total_failed', 0)
                total_emails = total_sent + total_failed
                start_datetime_str = log_data.get('timestamp_start', 'N/A')

                time_passed_text = "N/A"
                formatted_datetime = "N/A"
                if start_datetime_str != 'N/A':
                    try:
                        start_datetime_obj = datetime.datetime.strptime(start_datetime_str, "%Y-%m-%d %H:%M:%S")
                        now = datetime.datetime.now()
                        time_difference = now - start_datetime_obj
                        days = time_difference.days
                        hours, remainder = divmod(time_difference.seconds, 3600)
                        minutes, seconds = divmod(remainder, 60)
                        
                        if days > 0:
                            time_passed_text = f"{days} days ago"
                        elif hours > 0:
                            time_passed_text = f"{hours} hours ago"
                        elif minutes > 0:
                            time_passed_text = f"{minutes} minutes ago"
                        else:
                            time_passed_text = "Just now"

                        formatted_datetime = start_datetime_obj.strftime("%Y-%m-%d %I:%M %p")
                    except ValueError:
                        pass
                
                progress_text = ""
                if total_emails > 0:
                    progress_percent = int((total_sent / total_emails) * 100)
                    progress_text = f"{progress_percent}% ({total_sent}/{total_emails})"
                else:
                    progress_text = "0% (0/0)"

                self.analytics_tree.insert("", "end", iid=file, values=(campaign_name, total_sent, total_failed, formatted_datetime, time_passed_text))
        
        self.status_var.set("Status: Analytics table updated.")

    def show_campaign_details(self, log_data):
        self.clear_content()
        self.viewing_campaign_details_id = log_data.get('id')
        self.status_var.set(f"Status: Viewing details for '{log_data.get('name', 'campaign')}'.")
        ctk.CTkLabel(self.content_frame, text=f"Analytics for: {log_data.get('name')}", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)

        stats_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        stats_frame.pack(pady=10, fill="x", padx=20)
        stats_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="group1")

        sent_count = log_data.get('total_sent', 0)
        failed_count = log_data.get('total_failed', 0)
        total_count = sent_count + failed_count
        
        self.create_stat_card(stats_frame, "Total Emails", total_count, "#3498db", 0)
        self.create_stat_card(stats_frame, "Sent Successfully", sent_count, "#27ae60", 1)
        self.create_stat_card(stats_frame, "Failed", failed_count, "#e74c3c", 2)

        ctk.CTkLabel(self.content_frame, text="Detailed Email Log:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(20, 5), anchor="w", padx=20)
        
        search_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=20, pady=5)
        
        search_var = tk.StringVar()
        search_entry = ctk.CTkEntry(search_frame, textvariable=search_var, placeholder_text="Search emails...", width=300)
        search_entry.pack(side="left", padx=10)
        
        detail_frame = ctk.CTkFrame(self.content_frame)
        detail_frame.pack(fill="both", expand=True, padx=20)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2a2d2e", foreground="#dce4ee", fieldbackground="#2a2d2e", borderwidth=0)
        style.configure("Treeview.Heading", background="#343638", foreground="#dce4ee", font=("Arial", 12, "bold"))
        style.map('Treeview', background=[('selected', '#565b5e')])
        
        self.email_tree = ttk.Treeview(detail_frame, columns=("Recipient", "SMTP Used", "Status", "Reason", "Follow-up Status", "Follow-up Count"), show='headings')
        self.email_tree.heading("Recipient", text="Recipient")
        self.email_tree.heading("SMTP Used", text="SMTP Used")
        self.email_tree.heading("Status", text="Status")
        self.email_tree.heading("Reason", text="Reason")
        self.email_tree.heading("Follow-up Status", text="Follow-up Status")
        self.email_tree.heading("Follow-up Count", text="Follow-up Count")
        self.email_tree.column("Status", width=100)
        self.email_tree.column("Reason", width=150)
        self.email_tree.column("Follow-up Status", width=120)
        self.email_tree.column("Follow-up Count", width=120)
        self.email_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        
        scrollbar = ctk.CTkScrollbar(detail_frame, command=self.email_tree.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        self.email_tree.configure(yscrollcommand=scrollbar.set)
        
        search_entry.bind("<KeyRelease>", lambda event: self._filter_detailed_emails(search_var.get(), log_data))
        ctk.CTkButton(search_frame, text="Search", command=lambda: self._filter_detailed_emails(search_var.get(), log_data)).pack(side="left", padx=5)
        ctk.CTkButton(search_frame, text="Reset", command=lambda: self._filter_detailed_emails("", log_data)).pack(side="left", padx=5)

        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="Back to Analytics", command=self.show_analytics_ui).pack(side="left", padx=5)

        if self.email_tree:
            self._filter_detailed_emails("", log_data)

    def _filter_detailed_emails(self, query, log_data):
        if not self.email_tree: return
        
        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')

        for row in self.email_tree.get_children():
            self.email_tree.delete(row)
        
        for email_entry in log_data.get('emails', []):
            recipient = email_entry.get('recipient', '')
            if query.lower() in recipient.lower():
                followup_status = email_entry.get('followup_status', 'Not Sent')
                
                if recipient in self.blacklist_cache:
                    dnc_entry = self.blacklist_cache[recipient]
                    if dnc_entry.get('type') == 'lead':
                        followup_status = "Lead (DNC)"
                    else:
                        followup_status = "Blocklisted"
                elif email_entry.get('flag_no_followup'):
                        followup_status = "Manual Flag (No Follow-up)"
                
                row_values = (
                    recipient,
                    email_entry.get('smtp_used'),
                    email_entry.get('status'),
                    email_entry.get('reason'),
                    followup_status,
                    email_entry.get('followup_count', 0)
                )
                self.email_tree.insert("", "end", values=row_values)
    
    def show_follow_up_ui(self):
        self.clear_content()
        
        if self.followup_running:
            self.status_var.set(f"Follow-up in progress | Checked: {self.active_followup_info.get('checked', 0)}/{self.active_followup_info.get('total', 0)}")
            
            ctk.CTkLabel(self.content_frame, text="Live Follow-up Progress", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
            
            live_update_frame = ctk.CTkFrame(self.content_frame, fg_color="#2a2d2e", corner_radius=10)
            live_update_frame.pack(pady=20, padx=20, fill="x")
            
            ctk.CTkLabel(live_update_frame, textvariable=self.status_var, font=("Arial", 14, "bold")).pack(pady=5)
            
            self.followup_progress_bar = ctk.CTkProgressBar(live_update_frame, width=400)
            total = self.active_followup_info.get('total', 0)
            checked = self.active_followup_info.get('checked', 0)
            if total > 0:
                self.followup_progress_bar.set(checked / total)
            self.followup_progress_bar.pack(pady=10)
            
            def stop_followup_action():
                self.followup_running = False
                messagebox.showinfo("Follow-up Stopped", "The follow-up process is being stopped. Please wait.")
            
            button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
            button_frame.pack(pady=20)
            
            ctk.CTkButton(button_frame, text="Stop Follow-up", command=stop_followup_action, fg_color="#e74c3c", hover_color="#c0392b").pack(side="left", padx=10)
        else:
            self.status_var.set("Status: Ready to send follow-up campaigns.")
            ctk.CTkLabel(self.content_frame, text="Send Follow-up Campaign", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)

            filter_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
            filter_frame.pack(fill="x", padx=20, pady=10)
            
            search_var = tk.StringVar()
            search_entry = ctk.CTkEntry(filter_frame, textvariable=search_var, placeholder_text="Search campaigns...", width=200)
            search_entry.pack(side="left", padx=(0, 10))
            search_entry.bind("<KeyRelease>", lambda event: self._update_followup_campaign_list(search_var.get()))
            
            ctk.CTkLabel(filter_frame, text="Order by:").pack(side="left", padx=(10, 5))
            order_options = ["Newest First", "Oldest First"]
            order_var = ctk.StringVar(value=order_options[0])
            order_dropdown = ctk.CTkOptionMenu(filter_frame, values=order_options, variable=order_var, command=lambda x: self._update_followup_campaign_list(search_var.get(), order_var.get()))
            order_dropdown.pack(side="left", padx=5)

            campaign_list_frame = ctk.CTkFrame(self.content_frame)
            campaign_list_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview", background="#2a2d2e", foreground="#dce4ee", fieldbackground="#2a2d2e", borderwidth=0)
            style.configure("Treeview.Heading", background="#343638", foreground="#dce4ee", font=("Arial", 12, "bold"))
            style.map('Treeview', background=[('selected', '#565b5e')])
            
            self.followup_campaign_tree = ttk.Treeview(campaign_list_frame, columns=("Campaign Name", "Date", "Total Sent", "Unreplied"), show='headings')
            self.followup_campaign_tree.heading("Campaign Name", text="Campaign Name")
            self.followup_campaign_tree.heading("Date", text="Date")
            self.followup_campaign_tree.heading("Total Sent", text="Total Sent")
            self.followup_campaign_tree.heading("Unreplied", text="Unreplied")
            self.followup_campaign_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
            
            scrollbar = ctk.CTkScrollbar(campaign_list_frame, command=self.followup_campaign_tree.yview)
            scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
            self.followup_campaign_tree.configure(yscrollcommand=scrollbar.set)
            
            self._update_followup_campaign_list()

            def run_follow_up():
                selected_items = self.followup_campaign_tree.selection()
                if not selected_items:
                    messagebox.showerror("Error", "Please select a campaign.")
                    return
                
                campaign_id = selected_items[0]

                if self.all_campaign_logs.get(campaign_id) is None:
                    messagebox.showerror("Data Error", f"Could not find matching campaign data for ID: {campaign_id}")
                    return

                self.followup_thread = threading.Thread(
                    target=self._run_follow_up_campaign,
                    args=(campaign_id,),
                    daemon=True
                )
                self.followup_thread.start()
                
                campaign_name = self.all_campaign_logs[campaign_id].get('name', 'N/A')
                self.status_var.set(f"Status: Sending follow-ups for '{campaign_name}'.")

            ctk.CTkButton(self.content_frame, text="Send Follow-ups", command=run_follow_up, fg_color="#27ae60", hover_color="#2ecc71").pack(pady=10)
    
    def _update_followup_campaign_list(self, query="", order="Newest First"):
        if not self.followup_campaign_tree or not self.followup_campaign_tree.winfo_exists():
            return

        self.followup_campaign_tree.delete(*self.followup_campaign_tree.get_children())
        
        campaigns = list(self.all_campaign_logs.values())
        
        if order == "Newest First":
            campaigns.sort(key=lambda x: x.get('timestamp_start', '0'), reverse=True)
        else:
            campaigns.sort(key=lambda x: x.get('timestamp_start', '0'))
            
        filtered_campaigns = [c for c in campaigns if query.lower() in c.get('name', '').lower()]
        
        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')

        for campaign in filtered_campaigns:
            campaign_id = campaign.get('id')
            total_sent = campaign.get('total_sent', 0)
            
            unreplied_count = 0
            for email_entry in campaign.get('emails', []):
                is_blacklisted = email_entry.get('recipient') in self.blacklist_cache
                if email_entry.get('status') == 'sent' and email_entry.get('followup_status') != 'Replied' and not email_entry.get('flag_no_followup') and not is_blacklisted:
                    unreplied_count += 1
            
            start_datetime = campaign.get('timestamp_start', 'N/A')
            campaign_date = start_datetime.split(' ')[0] if ' ' in start_datetime else start_datetime
            
            self.followup_campaign_tree.insert("", "end", iid=campaign_id, values=(campaign['name'], campaign_date, total_sent, unreplied_count))
        
        self.status_var.set("Status: Follow-up campaign list updated.")

    # --- NEW: All methods for reply tracking and notifications ---
    def _reply_checker_loop(self):
        """Main loop for the background thread that periodically checks for replies."""
        while True:
            print("[REPLY CHECKER] Starting periodic check for new replies...")
            try:
                self._check_for_replies_background()
            except Exception as e:
                print(f"[REPLY CHECKER] An error occurred during the check: {e}")
            print(f"[REPLY CHECKER] Check finished. Waiting for {config.REPLY_CHECK_INTERVAL} seconds.")
            time.sleep(config.REPLY_CHECK_INTERVAL)

    def _check_for_replies_background(self):
        """Scans all campaigns for replies and sends notifications if new ones are found."""
        if not self.all_campaign_logs:
            print("[REPLY CHECKER] No campaign logs loaded yet. Skipping check.")
            return

        self.notifications_cache = self.load_json(config.NOTIFICATIONS_FILE, 'notifications_cache')
        notified_message_ids = {n['original_message_id'] for n in self.notifications_cache}
        
        emails_to_check = defaultdict(list)
        for campaign_log in self.all_campaign_logs.values():
            for email_entry in campaign_log.get('emails', []):
                if email_entry.get('status') == 'sent' and email_entry.get('message_id') and email_entry.get('message_id') not in notified_message_ids:
                    entry_with_campaign = email_entry.copy()
                    entry_with_campaign['campaign_name'] = campaign_log.get('name', 'N/A')
                    emails_to_check[email_entry['smtp_used']].append(entry_with_campaign)
        
        if not emails_to_check:
            print("[REPLY CHECKER] No new emails to check for replies.")
            return

        new_replies_found = False
        for smtp_email, entries in emails_to_check.items():
            smtp_account = self._get_smtp_account_by_email(smtp_email)
            if not smtp_account or not smtp_account.get('imap_server'):
                continue

            imap = None
            try:
                imap = imaplib.IMAP4_SSL(smtp_account['imap_server'])
                imap.login(smtp_account['email'], smtp_account['password'])
                imap.select('inbox')
                
                for entry in entries:
                    has_replied, reply_uid = self._has_replied_in_session(imap, entry)
                    if has_replied:
                        print(f"New reply detected from {entry['recipient']} for campaign '{entry['campaign_name']}'")
                        new_notification = {
                            "recipient": entry['recipient'],
                            "campaign_name": entry['campaign_name'],
                            "subject": entry['subject'],
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "original_message_id": entry['message_id'],
                            "seen": False
                        }
                        self.notifications_cache.append(new_notification)
                        self._send_admin_notification(new_notification)
                        new_replies_found = True
            except Exception as e:
                print(f"[REPLY CHECKER] IMAP error for {smtp_email}: {e}")
            finally:
                if imap:
                    imap.logout()

        if new_replies_found:
            self.save_json(config.NOTIFICATIONS_FILE, self.notifications_cache, 'notifications_cache')
            unread_count = sum(1 for n in self.notifications_cache if not n.get('seen', False))
            self.after(0, lambda: self.new_notifications_count.set(unread_count))


    def _send_admin_notification(self, notification_data):
        """Sends an email alert to the administrator about a new reply."""
        if not self.smtp_cache:
            print("[ADMIN NOTIFY] No SMTP accounts configured to send notification.")
            return

        admin_smtp = self.smtp_cache[0] # Use the first available SMTP
        subject = f"New Reply Received: {notification_data['recipient']}"
        body = (
            f"A new email reply has been detected.\n\n"
            f"From: {notification_data['recipient']}\n"
            f"Campaign: {notification_data['campaign_name']}\n"
            f"Original Subject: {notification_data['subject']}\n"
            f"Time: {notification_data['timestamp']}\n\n"
            f"Please check your inbox at <{notification_data['recipient']}> for the full message."
        )

        try:
            self.send_email(admin_smtp, config.ADMIN_EMAIL, subject, body)
            print(f"Admin notification sent successfully for reply from {notification_data['recipient']}")
        except Exception as e:
            print(f"Failed to send admin notification: {e}")
    
    def show_notifications_ui(self):
        """Displays the notification center UI."""
        self.clear_content()
        self.status_var.set("Status: Viewing notifications.")
        
        ctk.CTkLabel(self.content_frame, text="Notification Center", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=10)
        ctk.CTkLabel(self.content_frame, text="All detected replies are logged here, sorted by most recent.").pack(pady=(0, 20))

        if self.new_notifications_count.get() > 0:
            self.new_notifications_count.set(0)
            self.notifications_cache = self.load_json(config.NOTIFICATIONS_FILE, 'notifications_cache')
            for notification in self.notifications_cache:
                notification['seen'] = True
            self.save_json(config.NOTIFICATIONS_FILE, self.notifications_cache, 'notifications_cache')

        tree_frame = ctk.CTkFrame(self.content_frame)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)

        tree = ttk.Treeview(tree_frame, columns=("Time", "From", "Campaign", "Subject"), show='headings')
        tree.heading("Time", text="Time Detected")
        tree.heading("From", text="From")
        tree.heading("Campaign", text="Campaign")
        tree.heading("Subject", text="Original Subject")
        tree.column("Time", width=160)
        tree.column("From", width=200)
        tree.column("Campaign", width=200)

        tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        
        scrollbar = ctk.CTkScrollbar(tree_frame, command=tree.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        tree.configure(yscrollcommand=scrollbar.set)
        
        sorted_notifications = sorted(self.notifications_cache, key=lambda x: x['timestamp'], reverse=True)
        
        for notification in sorted_notifications:
            tree.insert("", "end", values=(
                notification['timestamp'],
                notification['recipient'],
                notification['campaign_name'],
                notification['subject']
            ))

    # --- NEW: DNC Management UI and Logic ---
    def show_dnc_ui(self):
        self.clear_content()
        self.status_var.set("Status: Managing Do-Not-Contact list.")
        
        ctk.CTkLabel(self.content_frame, text="Do-Not-Contact (DNC) Management", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=10)
        
        # Main frame splitting UI into two parts
        main_dnc_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        main_dnc_frame.pack(fill="both", expand=True, padx=20, pady=10)
        main_dnc_frame.grid_columnconfigure(0, weight=1)
        main_dnc_frame.grid_columnconfigure(1, weight=2)
        main_dnc_frame.grid_rowconfigure(0, weight=1)

        # Left frame for adding entries
        add_frame = ctk.CTkFrame(main_dnc_frame)
        add_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        tab_view = ctk.CTkTabview(add_frame)
        tab_view.pack(fill="both", expand=True, padx=10, pady=10)
        tab_view.add("Add Leads")
        tab_view.add("Add to Blocklist")

        # Leads Tab
        ctk.CTkLabel(tab_view.tab("Add Leads"), text="Paste Lead Emails (one per line):").pack(anchor="w", padx=10, pady=(10,0))
        leads_textbox = ctk.CTkTextbox(tab_view.tab("Add Leads"), height=150)
        leads_textbox.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(tab_view.tab("Add Leads"), text="Comment (optional):").pack(anchor="w", padx=10, pady=(10,0))
        leads_comment_entry = ctk.CTkEntry(tab_view.tab("Add Leads"), placeholder_text="e.g., Contacted via phone")
        leads_comment_entry.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(tab_view.tab("Add Leads"), text="Add Leads to DNC", 
                      command=lambda: self._add_to_dnc_list(leads_textbox.get("1.0", "end-1c"), 'lead', leads_comment_entry.get())).pack(pady=20)
        
        # Blocklist Tab
        ctk.CTkLabel(tab_view.tab("Add to Blocklist"), text="Paste Emails to Block (one per line):").pack(anchor="w", padx=10, pady=(10,0))
        blocklist_textbox = ctk.CTkTextbox(tab_view.tab("Add to Blocklist"), height=150)
        blocklist_textbox.pack(fill="x", expand=True, padx=10, pady=5)
        
        ctk.CTkButton(tab_view.tab("Add to Blocklist"), text="Add to Blocklist", 
                      command=lambda: self._add_to_dnc_list(blocklist_textbox.get("1.0", "end-1c"), 'blocklist')).pack(pady=20)
        
        # Right frame for displaying the list
        list_frame = ctk.CTkFrame(main_dnc_frame)
        list_frame.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(list_frame, text="Current DNC List", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        tree_container = ctk.CTkFrame(list_frame)
        tree_container.pack(fill="both", expand=True, padx=10, pady=5)

        self.dnc_tree = ttk.Treeview(tree_container, columns=("Email", "Type", "Comment", "Date Added"), show='headings')
        self.dnc_tree.heading("Email", text="Email")
        self.dnc_tree.heading("Type", text="Type")
        self.dnc_tree.heading("Comment", text="Comment")
        self.dnc_tree.heading("Date Added", text="Date Added")
        self.dnc_tree.column("Type", width=80, anchor="center")
        self.dnc_tree.column("Date Added", width=150)
        self.dnc_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ctk.CTkScrollbar(tree_container, command=self.dnc_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.dnc_tree.configure(yscrollcommand=scrollbar.set)
        
        ctk.CTkButton(list_frame, text="Remove Selected", command=self._remove_from_dnc, fg_color="#e74c3c", hover_color="#c0392b").pack(pady=10)
        
        self._populate_dnc_tree()

    def _populate_dnc_tree(self):
        if not hasattr(self, 'dnc_tree') or not self.dnc_tree.winfo_exists():
            return
            
        for row in self.dnc_tree.get_children():
            self.dnc_tree.delete(row)

        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')
        
        # Sort by date added, newest first
        sorted_dnc = sorted(self.blacklist_cache.items(), key=lambda item: item[1].get('date_added', '0'), reverse=True)
        
        for email, details in sorted_dnc:
            self.dnc_tree.insert("", "end", iid=email, values=(
                email,
                details.get('type', 'N/A').title(),
                details.get('comment', ''),
                details.get('date_added', 'N/A')
            ))

    def _add_to_dnc_list(self, emails_text, type, comment=""):
        emails_to_add = [email.strip() for email in emails_text.splitlines() if email.strip() and '@' in email]
        if not emails_to_add:
            messagebox.showerror("Error", "No valid email addresses were provided.")
            return

        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        for email in emails_to_add:
            self.blacklist_cache[email] = {
                "type": type,
                "comment": comment,
                "date_added": now
            }
        
        self.save_json(config.BLACKLIST_FILE, self.blacklist_cache, 'blacklist_cache')
        
        if type == 'lead':
            # Run this in a thread to avoid freezing UI if logs are large
            threading.Thread(target=self._update_logs_for_new_dnc, args=(emails_to_add,), daemon=True).start()
        
        self.status_var.set(f"Status: Added {len(emails_to_add)} email(s) to DNC list.")
        self._populate_dnc_tree()
        messagebox.showinfo("Success", f"Successfully added {len(emails_to_add)} email(s) to the DNC list.")

    def _remove_from_dnc(self):
        selected_items = self.dnc_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "Please select one or more emails to remove.")
            return

        if not messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove {len(selected_items)} email(s) from the DNC list?"):
            return

        self.blacklist_cache = self.load_json(config.BLACKLIST_FILE, 'blacklist_cache')
        
        removed_count = 0
        for email in selected_items:
            if email in self.blacklist_cache:
                del self.blacklist_cache[email]
                removed_count += 1
        
        self.save_json(config.BLACKLIST_FILE, self.blacklist_cache, 'blacklist_cache')
        self.status_var.set(f"Status: Removed {removed_count} email(s) from DNC list.")
        self._populate_dnc_tree()

    def _update_logs_for_new_dnc(self, emails_to_flag):
        self.after(0, lambda: self.status_var.set("Status: Updating past campaign logs... Please wait."))
        
        logs_to_resave = set()
        for log_file, log_data in self.all_campaign_logs.items():
            was_modified = False
            for email_entry in log_data.get('emails', []):
                if email_entry.get('recipient') in emails_to_flag:
                    if not email_entry.get('flag_no_followup'):
                        email_entry['flag_no_followup'] = True
                        was_modified = True
            
            if was_modified:
                logs_to_resave.add(log_file)
        
        if logs_to_resave:
            for log_file in logs_to_resave:
                self.save_json(os.path.join(config.LOG_DIR, log_file), self.all_campaign_logs[log_file])
            self.after(0, lambda: self.status_var.set(f"Status: Finished updating {len(logs_to_resave)} campaign logs."))
        else:
            self.after(0, lambda: self.status_var.set("Status: No past campaign logs needed updates."))


# ------------------------- 4. Run Application ------------------------ #
if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()

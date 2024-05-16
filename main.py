import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import messagebox
import requests
import time
from datetime import datetime, timedelta
from tkinter import filedialog
import os
import csv
from dateutil import parser

class YouTubeScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("YouTube Scrapper")

        self.root.resizable(False, False)

        self.root.set_theme("black")

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        window_width = int(screen_width / 2)
        window_height = int(screen_height / 2)
        x_coordinate = int((screen_width - window_width) / 2)
        y_coordinate = int((screen_height - window_height) / 2)

        self.root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

        self.tabControl = ttk.Notebook(self.root)

        self.tab2 = ttk.Frame(self.tabControl)
        self.tab3 = ttk.Frame(self.tabControl)

        self.tabControl.add(self.tab2, text="Views During a Period")
        self.tabControl.add(self.tab3, text="Get All Data")
        self.tabControl.bind("<<NotebookTabChanged>>", self.show_message)
        self.tabControl.pack(expand=1, fill="both")

        self.channel_url_var = tk.StringVar()
        self.date1_var = tk.StringVar()
        self.date2_var = tk.StringVar()

        self.dates_entered = False

        self.views_label = tk.Label(self.tab2, text="Views During Period: N/A", bg="grey", font=("Helvetica", 14))
        self.views_label.grid(row=4, column=0, columnspan=2, pady=(20, 5), sticky='nsew')

        self.likes_label = tk.Label(self.tab2, text="Total Likes During Period: N/A", bg="grey", font=("Helvetica", 14))
        self.likes_label.grid(row=5, column=0, columnspan=2, pady=(5, 5), sticky='nsew')

        self.comments_label = tk.Label(self.tab2, text="Total Comments During Period: N/A", bg="grey", font=("Helvetica", 14))
        self.comments_label.grid(row=6, column=0, columnspan=2, pady=(5, 20), sticky='nsew')

        label_channel_url_tab2 = tk.Label(self.tab2, text="Enter Channel URL:", bg="grey", font=("Helvetica", 12))
        label_channel_url_tab2.grid(row=0, column=0, padx=10, pady=10, sticky='e')
        entry_channel_url_tab2 = tk.Entry(self.tab2, textvariable=self.channel_url_var, font=("Helvetica", 12), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0, bg="white", fg="black", insertbackground="black", selectbackground="#a6a6a6", selectforeground="black")
        entry_channel_url_tab2.grid(row=0, column=1, padx=10, pady=10, sticky='w')
        label_start_date_tab2 = tk.Label(self.tab2, text="Enter Start Date (YYYY-MM-DD):", bg="grey", font=("Helvetica", 12))
        label_start_date_tab2.grid(row=1, column=0, padx=10, pady=10, sticky='e')
        entry_start_date_tab2 = tk.Entry(self.tab2, textvariable=self.date1_var, font=("Helvetica", 12), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0, bg="white", fg="black", insertbackground="black", selectbackground="#a6a6a6", selectforeground="black")
        entry_start_date_tab2.grid(row=1, column=1, padx=10, pady=10, sticky='w')
        label_end_date_tab2 = tk.Label(self.tab2, text="Enter End Date (YYYY-MM-DD):", bg="grey", font=("Helvetica", 12))
        label_end_date_tab2.grid(row=2, column=0, padx=10, pady=10, sticky='e')
        entry_end_date_tab2 = tk.Entry(self.tab2, textvariable=self.date2_var, font=("Helvetica", 12), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0, bg="white", fg="black", insertbackground="black", selectbackground="#a6a6a6", selectforeground="black")
        entry_end_date_tab2.grid(row=2, column=1, padx=10, pady=10, sticky='w')
        button_run_tab2 = tk.Button(self.tab2, text="Run", command=self.run_views_during_period, bg="red", fg="black", font=("Helvetica", 12, "bold"), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0)
        button_run_tab2.grid(row=3, column=0, columnspan=2, pady=10)

        self.tab2.grid_columnconfigure(0, weight=1)
        self.tab2.grid_columnconfigure(1, weight=1)
        self.tab2.grid_rowconfigure(0, weight=1)
        self.tab2.grid_rowconfigure(1, weight=1)
        self.tab2.grid_rowconfigure(2, weight=1)
        self.tab2.grid_rowconfigure(3, weight=1)
        self.tab2.grid_rowconfigure(4, weight=1)

        label_channel_url_tab3 = tk.Label(self.tab3, text="Enter Channel URL:", bg="grey", font=("Helvetica", 12))
        label_channel_url_tab3.grid(row=0, column=0, padx=10, pady=10, sticky='e')
        entry_channel_url_tab3 = tk.Entry(self.tab3, textvariable=self.channel_url_var, font=("Helvetica", 12), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0, bg="white", fg="black", insertbackground="black", selectbackground="#a6a6a6", selectforeground="black")
        entry_channel_url_tab3.grid(row=0, column=1, padx=10, pady=10, sticky='w')
        label_start_date_tab3 = tk.Label(self.tab3, text="Enter Start Date (YYYY-MM-DD):", bg="grey", font=("Helvetica", 12))
        label_start_date_tab3.grid(row=1, column=0, padx=10, pady=10, sticky='e')
        entry_start_date_tab3 = tk.Entry(self.tab3, textvariable=self.date1_var, font=("Helvetica", 12), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0, bg="white", fg="black", insertbackground="black", selectbackground="#a6a6a6", selectforeground="black")
        entry_start_date_tab3.grid(row=1, column=1, padx=10, pady=10, sticky='w')
        label_end_date_tab3 = tk.Label(self.tab3, text="Enter End Date (YYYY-MM-DD):", bg="grey", font=("Helvetica", 12))
        label_end_date_tab3.grid(row=2, column=0, padx=10, pady=10, sticky='e')
        entry_end_date_tab3 = tk.Entry(self.tab3, textvariable=self.date2_var, font=("Helvetica", 12), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0, bg="white", fg="black", insertbackground="black", selectbackground="#a6a6a6", selectforeground="black")
        entry_end_date_tab3.grid(row=2, column=1, padx=10, pady=10, sticky='w')
        button_run_tab3 = tk.Button(self.tab3, text="Run", command=self.run_get_all_data, bg="red", fg="black", font=("Helvetica", 12, "bold"), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0)
        button_run_tab3.grid(row=3, column=0, columnspan=2, pady=10)

        self.tab3.grid_columnconfigure(0, weight=1)
        self.tab3.grid_columnconfigure(1, weight=1)
        self.tab3.grid_rowconfigure(0, weight=1)
        self.tab3.grid_rowconfigure(1, weight=1)
        self.tab3.grid_rowconfigure(2, weight=1)
        self.tab3.grid_rowconfigure(3, weight=1)
        self.tab3.grid_rowconfigure(4, weight=1)

        self.root.mainloop()

    def show_message(self, event):
        current_tab = self.tabControl.index(self.tabControl.select())
        if current_tab == 1: 
            messagebox.showinfo("Message", "To import URLs from text file, leave URL field empty and click 'RUN'.")

    def run_views_during_period(self):
        url = self.channel_url_var.get()
        date1 = self.date1_var.get()
        date2 = self.date2_var.get()

        func_call_date1 = date1
        func_call_date2 = date2

        if not date1 or not date2:
            messagebox.showinfo("Error", "Please enter both start and end dates.")
            return

        if date1 > date2:
            messagebox.showinfo("Error", "Invalid date range.")
            return

        current_date = datetime.now().date()

        if date2:
            date2_datetime = datetime.fromisoformat(date2).date()
            delta = current_date - date2_datetime

            if timedelta(days=30) >= delta > timedelta(days=6*30):
                max_results = 50
                maxShortResults = 50

            elif delta > timedelta(days=6*30):
                max_results = 200
                maxShortResults = 200

            else:
                maxShortResults = 10
                max_results = 7
        else:
            messagebox.showinfo("Error", "Invalid date range.")
            max_results = None 

        input_data = {
            "searchKeywords": url,
            "startUrls": [{"url": url}],
            "maxResults": max_results,
            "maxResultsShorts": maxShortResults
        }
        result = self.run_actor(input_data)

        date_list = sorted(set(item['date'] for item in result))
        date_list = [datetime.fromisoformat(date.replace('Z', '')).date() for date in date_list]

        date1 = datetime.fromisoformat(date1).date() if date1 else None
        date2 = datetime.fromisoformat(date2).date() if date2 else None

        for i in range(len(date_list)):
            if date_list[i] >= date1:
                start_date = date_list[i]
                break
            else:
                start_date = None
        
        if start_date == None:
            messagebox.showinfo("Error", "Invalid date range or no data available.")
            return

        for i in range(len(date_list)):
            if date_list[i] >= date2 and i == 0:
                messagebox.showinfo("Error", "Invalid date range or no data available.")
                return
            elif date_list[i] >= date2 and i > 0:
                end_date = date_list[i-1]
                break
            else:
                end_date = None
        
        if end_date == None:
            end_date = max(date_list) if date_list else None

        if start_date is None or end_date is None:
            messagebox.showinfo("Error", "Invalid date range or no data available.")
            return
        
        if start_date == end_date:
            end_date = start_date + timedelta(days=1)

        if not start_date or not end_date or start_date > end_date:
            messagebox.showinfo("Error", "Invalid date range or no data available.")
            return

        start_date_str = start_date.strftime('%Y-%m-%dT%H:%M:%S%z')
        end_date_str = end_date.strftime('%Y-%m-%dT%H:%M:%S%z')

        filtered_result = [item for item in result if start_date_str <= item['date'] <= end_date_str]

        first_element = filtered_result[0]
        channel_name = first_element.get('channelName', 'N/A')

        subscribers_count = result[0].get('numberOfSubscribers', 0)
        views_sum = sum(item['viewCount'] for item in filtered_result if item['viewCount'] is not None)
        likes_sum = sum(item['likes'] if item['likes'] is not None else 0 for item in filtered_result)
        comments_sum = sum(item['commentsCount'] for item in filtered_result if item['commentsCount'] is not None)

        result_text = f"Views During {date1} to {date2}: {views_sum}"
        self.views_label.config(text=result_text)
        likes_text = f"Total Likes During {date1} to {date2}: {likes_sum}"
        self.likes_label.config(text=likes_text)
        comments_text = f"Total Comments During {date1} to {date2}: {comments_sum}"
        self.comments_label.config(text=comments_text)
        self.root.update()

        self.button_export_txt = tk.Button(self.tab2, text="Export to txt", command=lambda: self.export_to_txt(func_call_date1, func_call_date2, channel_name, subscribers_count, likes_sum, views_sum, comments_sum), bg="grey", fg="black", font=("Helvetica", 12, "bold"), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0)
        self.button_export_txt.grid(row=3, column=1, columnspan=2, pady=4)

    def export_to_txt(self, date1, date2, channelName, subscribersCount, totalLikes, totalViews, totalComments):
        new_format_date1 = datetime.strptime(date1, '%Y-%m-%d').strftime('%Y %m %d')
        new_format_date2 = datetime.strptime(date2, '%Y-%m-%d').strftime('%Y %m %d')

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        youtube_data_path = os.path.join(desktop_path, "Youtube Data")
        if not os.path.exists(youtube_data_path):
            os.makedirs(youtube_data_path)

        channel_folder_path = os.path.join(youtube_data_path, channelName)
        if not os.path.exists(channel_folder_path):
            os.makedirs(channel_folder_path)

        filename = f"{channelName}_{new_format_date1}_{new_format_date2}.txt"
        file_path = os.path.join(channel_folder_path, filename)

        with open(file_path, 'w') as text_file:
            text_file.write(f"Channel Name: {channelName}\n")
            text_file.write(f"Subscriber Count: {subscribersCount}\n")
            text_file.write(f"Date Range: {new_format_date1} - {new_format_date2}\n")
            text_file.write(f"Total Likes: {totalLikes}\n")
            text_file.write(f"Total Views: {totalViews}\n")
            text_file.write(f"Total Comments: {totalComments}\n")
        
        messagebox.showinfo("Info", "Text file exported.")

    def run_get_all_data(self):
        url = self.channel_url_var.get()
        date1 = self.date1_var.get()
        date2 = self.date2_var.get()

        if not url:
            self.button_import_from_txt = tk.Button(self.tab3, text="Import from TXT", command=lambda: self.import_from_txt(), bg="grey", fg="black", font=("Helvetica", 12, "bold"), bd=2, relief=tk.GROOVE, borderwidth=2, highlightthickness=0)
            self.button_import_from_txt.grid(row=4, column=0, columnspan=2, pady=10)
        
        if not date1 or not date2:
            messagebox.showinfo("Error", "Please enter both start and end dates.")
            return

        if date1 > date2:
            messagebox.showinfo("Error", "Start date should be earlier than End date.")
            return

        current_date = datetime.now().date()

        if date2:
            date2_datetime = datetime.fromisoformat(date2).date()
            delta = current_date - date2_datetime

            if timedelta(days=30) >= delta > timedelta(days=6*30):
                max_results = 50
                maxShortResults = 50

            elif delta > timedelta(days=6*30):
                max_results = 200
                maxShortResults = 200

            else:
                maxShortResults = 10
                max_results = 7
        else:
            messagebox.showinfo("Error", "Invalid date range.")
            max_results = None 

        input_data = {
            "searchKeywords": url,
            "startUrls": [{"url": url}],
            "maxResults": max_results,
            "maxResultsShorts": maxShortResults
        }
        result = self.run_actor(input_data)

        date_list = sorted(set(item['date'] for item in result))
        date_list = [datetime.fromisoformat(date.replace('Z', '')).date() for date in date_list]

        date1 = datetime.fromisoformat(date1).date() if date1 else None
        date2 = datetime.fromisoformat(date2).date() if date2 else None

        for i in range(len(date_list)):
            if date_list[i] >= date1:
                start_date = date_list[i]
                break
            else:
                start_date = None
        
        if start_date == None:
            messagebox.showinfo("Error", "Invalid date range or no data available.")
            return

        for i in range(len(date_list)):
            if date_list[i] >= date2 and i == 0:
                messagebox.showinfo("Error", "Invalid date range or no data available.")
                return
            elif date_list[i] >= date2 and i > 0:
                end_date = date_list[i-1]
                break
            else:
                end_date = None
        
        if end_date == None:
            end_date = max(date_list) if date_list else None

        if start_date is None or end_date is None:
            messagebox.showinfo("Error", "Invalid date range or no data available.")
            return
        
        end_date = end_date + timedelta(days=1)

        if not start_date or not end_date or start_date > end_date:
            messagebox.showinfo("Error", "Invalid date range or no data available.")
            return

        start_date_str = start_date.strftime('%Y-%m-%dT%H:%M:%S%z')
        
        end_date_str = end_date.strftime('%Y-%m-%dT%H:%M:%S%z')

        filtered_result = [item for item in result if start_date_str <= item['date'] <= end_date_str]

        self.show_result_table(filtered_result)
        self.root.update()

    def run_actor(self, input_data):
        apify_token = '' # Use your own API key here
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {apify_token}',
        }

        actor_response = requests.post('https://api.apify.com/v2/acts/streamers~youtube-scraper/runs?token=', json=input_data, headers=headers) # place API token/Key after the '=' sign

        if actor_response.status_code == 201:
            actor_run_info = actor_response.json()
            actor_run_id = actor_run_info['data']['id']
            print(f"Actor run initiated. Waiting for completion...")

            while True:
                run_status_response = requests.get(f'https://api.apify.com/v2/actor-runs/{actor_run_id}', headers=headers)
                run_status_info = run_status_response.json()
                run_status = run_status_info['data']['status']
                print(f"Actor run status: {run_status}. Waiting for completion...")

                if run_status in ['SUCCEEDED', 'FAILED', 'ABORTED']:
                    break

                time.sleep(30)

            if run_status == 'SUCCEEDED':
                dataset_id = run_status_info['data']['defaultDatasetId']
                dataset_url = f'https://api.apify.com/v2/datasets/{dataset_id}/items'
                dataset_response = requests.get(dataset_url, headers=headers)

                if dataset_response.status_code == 200:
                    dataset_items = dataset_response.json()
                    return dataset_items
                else:
                    print(f"Error fetching dataset items: {dataset_response.status_code} - {dataset_response.text}")
            else:
                print(f"Actor run did not succeed. Status: {run_status}")
        else:
            print(f"Error starting actor run: {actor_response.status_code} - {actor_response.text}")

    def show_result_table(self, result):
        result_window = tk.Toplevel(self.root)
        result_window.title("Summarized Results")

        columns = [
            "Title", "URL", "View Count", "Date", "Like", "Channel Name",
            "Channel Url", "Channel Join Date", "Channel Location",
            "Channel Total Videos", "Channel Total Views", "Number of Subscribers",
            "Duration", "Comments Count", "Comments Turned Off", "Is Monetized"
        ]

        columns_json_mapping = {
            "Title": "title", "URL": "url", "View Count": "viewCount", "Date": "date", "Like": "likes", "Channel Name": "channelName",
            "Channel Url": "channelUrl", "Channel Join Date": "channelJoinedDate", "Channel Location": "channelLocation",
            "Channel Total Videos": "channelTotalVideos", "Channel Total Views": "channelTotalViews", "Number of Subscribers": "numberOfSubscribers",
            "Duration": "duration", "Comments Count": "commentsCount", "Comments Turned Off": "commentsTurnedOff", "Is Monetized": "isMonetized"
        }

        table = ttk.Treeview(
            result_window,
            columns=columns,
            show='headings',
            selectmode='browse',
        )

        for col in columns:
            table.heading(col, text=col.capitalize())
            table.column(col, width=150, minwidth=50, stretch=tk.NO) 

        table.bind("<Button-3>", lambda e: self.show_context_menu(e, table))

        for item in result:
            row_data = [item.get(columns_json_mapping[col], "") for col in columns]
            table.insert("", "end", values=row_data)

        table.pack(expand=True, fill='both')

        y_scrollbar = ttk.Scrollbar(result_window, orient='vertical', command=table.yview)
        y_scrollbar.pack(side='right', fill='y')

        table.configure(yscrollcommand=y_scrollbar.set)

        x_scrollbar = ttk.Scrollbar(result_window, orient='horizontal', command=table.xview)
        x_scrollbar.pack(side='bottom', fill='x')

        table.configure(xscrollcommand=x_scrollbar.set)

        show_detailed_button = tk.Button(result_window, text="See All Detailed Data", command=lambda: self.show_detailed_table(result))
        show_detailed_button.pack(side='bottom', pady=10)

        result_window.state("zoomed")

        table.bind("<Configure>", lambda e: self.update_column_width(table))
        self.check_for_events()

    def check_for_events(self):
        self.root.update()
        self.root.after(100, self.check_for_events)
        
    def update_column_width(self, table):
        for col in table["columns"]:
            max_width = max(
                table.bbox(item, col)[2] - table.bbox(item, col)[0] + 10
                for item in table.get_children('')
            )
            table.column(col, width=max_width)

    def show_context_menu(self, event, table):
        context_menu = tk.Menu(table, tearoff=0)
        context_menu.add_command(label="Copy", command=lambda: self.copy_selected_data(table))
        context_menu.post(event.x_root, event.y_root)

    def copy_selected_data(self, table):
        selected_item = table.selection()
        if selected_item:
            selected_data = table.item(selected_item, 'values')
            self.root.clipboard_clear()
            self.root.clipboard_append('\t'.join(selected_data))
            self.root.update()

    def show_detailed_table(self, result):
        detailed_window = tk.Toplevel(self.root)
        detailed_window.title("Detailed Results")

        columns_detailed = [
            "title", "id", "url", "thumbnailUrl", "viewCount", "date", "likes",
            "location", "channelName", "channelUrl", "channelDescription", "channelJoinedDate",
            "channelLocation", "channelTotalVideos", "channelTotalViews", "numberOfSubscribers",
            "inputChannelUrl", "duration", "commentsCount", "text", "subtitles", "commentsTurnedOff",
            "comments", "fromYTUrl", "isMonetized"
        ]

        table_detailed = ttk.Treeview(
            detailed_window,
            columns=columns_detailed,
            show='headings',
            selectmode='browse',
        )

        for col in columns_detailed:
            table_detailed.heading(col, text=col.capitalize()) 
            table_detailed.column(col, width=150, minwidth=50, stretch=tk.NO) 
            
        for item in result:
            row_data_detailed = [item.get(col, "") for col in columns_detailed]
            table_detailed.insert("", "end", values=row_data_detailed)

        table_detailed.pack(expand=True, fill='both')

        y_scrollbar_detailed = ttk.Scrollbar(detailed_window, orient='vertical', command=table_detailed.yview)
        y_scrollbar_detailed.pack(side='right', fill='y')

        table_detailed.configure(yscrollcommand=y_scrollbar_detailed.set)

        x_scrollbar_detailed = ttk.Scrollbar(detailed_window, orient='horizontal', command=table_detailed.xview)
        x_scrollbar_detailed.pack(side='bottom', fill='x')

        table_detailed.configure(xscrollcommand=x_scrollbar_detailed.set)

        detailed_window.state("zoomed")

        table_detailed.bind("<Configure>", lambda e: self.update_column_width(table_detailed))

        export_button = tk.Button(detailed_window, text="Export to CSV", command=lambda: self.export_to_csv(table_detailed))
        export_button.pack(side='bottom', pady=10)

    def import_from_txt(self):
        txt_file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

        if not txt_file_path:
            messagebox.showinfo("Info", "No TXT file selected.")
            return

        with open(txt_file_path, 'r') as txt_file:
            urls = txt_file.read().splitlines()

        if not urls:
            messagebox.showinfo("Info", "No URLs found in the selected TXT file.")
            return

        for url in urls:
            date1 = self.date1_var.get()
            date2 = self.date2_var.get()
            
            if not date1 or not date2:
                messagebox.showinfo("Error", "Please enter both start and end dates.")
                return

            if date1 > date2:
                messagebox.showinfo("Error", "Start date should be earlier than End date.")
                return

            current_date = datetime.now().date()

            if date2:
                date2_datetime = datetime.fromisoformat(date2).date()
                delta = current_date - date2_datetime

                if timedelta(days=30) >= delta > timedelta(days=6*30):
                    max_results = 50
                    maxShortResults = 50

                elif delta > timedelta(days=6*30):
                    max_results = 200
                    maxShortResults = 200

                else:
                    maxShortResults = 10
                    max_results = 7
            else:
                messagebox.showinfo("Error", "Invalid date range.")
                max_results = None 

            input_data = {
                "searchKeywords": url,
                "startUrls": [{"url": url}],
                "maxResults": max_results,
                "maxResultsShorts": maxShortResults
            }
            result = self.run_actor(input_data)

            date_list = sorted(set(item['date'] for item in result))
            date_list = [datetime.fromisoformat(date.replace('Z', '')).date() for date in date_list]

            date1 = datetime.fromisoformat(date1).date() if date1 else None
            date2 = datetime.fromisoformat(date2).date() if date2 else None

            for i in range(len(date_list)):
                if date_list[i] >= date1:
                    start_date = date_list[i]
                    break
                else:
                    start_date = None
            
            if start_date == None:
                messagebox.showinfo("Error", "Invalid date range or no data available.")
                return

            for i in range(len(date_list)):
                if date_list[i] >= date2 and i == 0:
                    messagebox.showinfo("Error", "Invalid date range or no data available.")
                    return
                elif date_list[i] >= date2 and i > 0:
                    end_date = date_list[i-1]
                    break
                else:
                    end_date = None
            
            if end_date == None:
                end_date = max(date_list) if date_list else None

            if start_date is None or end_date is None:
                messagebox.showinfo("Error", "Invalid date range or no data available.")
                return
            
            end_date = end_date + timedelta(days=1)

            if not start_date or not end_date or start_date > end_date:
                messagebox.showinfo("Error", "Invalid date range or no data available.")
                return

            start_date_str = start_date.strftime('%Y-%m-%dT%H:%M:%S%z')
            
            end_date_str = end_date.strftime('%Y-%m-%dT%H:%M:%S%z')

            filtered_result = [item for item in result if start_date_str <= item['date'] <= end_date_str]

            channel_name = result[0].get("channelName", "Unknown")

            columns_detailed = [
                "title", "id", "url", "thumbnailUrl", "viewCount", "date", "likes",
                "location", "channelName", "channelUrl", "channelDescription", "channelJoinedDate",
                "channelLocation", "channelTotalVideos", "channelTotalViews", "numberOfSubscribers",
                "inputChannelUrl", "duration", "commentsCount", "text", "subtitles", "commentsTurnedOff",
                "comments", "fromYTUrl", "isMonetized"
            ]

            table_detailed = ttk.Treeview(columns=columns_detailed, show='headings', selectmode='browse')

            for col in columns_detailed:
                table_detailed.heading(col, text=col.capitalize())
                table_detailed.column(col, width=150, minwidth=50, stretch=tk.NO)

            for item in filtered_result:
                row_data_detailed = [item.get(col, "") for col in columns_detailed]
                table_detailed.insert("", "end", values=row_data_detailed)

            self.export_to_csv(table_detailed)

        messagebox.showinfo("Info", "Data exported from TXT file to Desktop\Youtube Data")

    def export_to_csv(self, table):
        columns = table["columns"]
        data = [table.item(item, 'values') for item in table.get_children('')]

        if not data:
            messagebox.showinfo("Info", "No data to export.")
            return

        first_row = table.item(table.get_children('')[0], 'values')
        channel_name = first_row[columns.index("channelName")]

        date_column_index = columns.index("date")
        start_date_str = data[0][date_column_index]
        end_date_str = data[-1][date_column_index]

        try:
            start_date = parser.isoparse(start_date_str)
            end_date = parser.isoparse(end_date_str)
        except ValueError as ve:
            messagebox.showinfo("Error", f"Error parsing date: {str(ve)}")
            return

        base_folder = os.path.join(os.path.expanduser("~"), "Desktop", "Youtube Data")
        os.makedirs(base_folder, exist_ok=True)

        channel_folder = os.path.join(base_folder, channel_name)
        os.makedirs(channel_folder, exist_ok=True)

        csv_file_name = f"{channel_name} {start_date.strftime('%Y %m %d')}_{end_date.strftime('%Y %m %d')}.csv"

        file_path = os.path.join(channel_folder, csv_file_name)

        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow(columns)
                csv_writer.writerows(data)
            messagebox.showinfo("Info", f"Data exported to:\n{file_path}")
        except Exception as e:
            messagebox.showinfo("Error", f"Error exporting data: {str(e)}")


if __name__ == "__main__":
    root = ThemedTk(theme="black") 
    app = YouTubeScraperGUI(root)
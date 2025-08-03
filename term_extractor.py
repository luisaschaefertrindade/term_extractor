import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import spacy
from collections import Counter
import os
import pandas as pd
import re
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
from docx import Document
from spacy.matcher import Matcher
import threading

# --- Your extraction utilities ---

def chunk_text(text, max_chars=200_000):
    chunks = []
    start = 0
    text_len = len(text)
    while start < text_len:
        end = start + max_chars
        boundary = text.rfind("\n\n", start, end)
        if boundary == -1 or boundary <= start:
            boundary = end
        chunk = text[start:boundary]
        chunks.append(chunk.strip())
        start = boundary
    return chunks

def extract_terms_with_context(chunks, nlp_model, progress_callback=None):
    term_counter = Counter()
    term_contexts = {}
    matcher = Matcher(nlp_model.vocab)
    patterns = [
        [{"POS": "ADJ", "OP": "*"}, {"POS": "NOUN", "OP": "+"}],                            # Adjective + Noun(s)
        [{"POS": "NOUN"}, {"LOWER": "of"}, {"POS": "NOUN"}],                                # Noun of Noun
        [{"POS": "NOUN"}, {"POS": "NOUN"}],                                                 # Noun + Noun
        [{"POS": "ADJ"}, {"POS": "NOUN"}, {"POS": "NOUN"}],                                  # Adjective + Noun + Noun
        [{"POS": "PROPN", "OP": "+"}],                                                      # Named entities (PROPN chunks)
    ]

    matcher.add("CANDIDATE_TERMS", patterns)
    for i, chunk in enumerate(chunks):
        doc = nlp_model(chunk)
        matches = matcher(doc)
        for match_id, start, end in matches:
            span = doc[start:end]
            term = span.text.strip().lower()
            if len(term) > 1:
                term_counter[term] += 1
                # Store ALL contexts per term as a list
                if term not in term_contexts:
                    term_contexts[term] = []
                context_sentence = span.sent.text.strip() if span.sent else ""
                term_contexts[term].append(context_sentence)
        if progress_callback:
            progress_callback(i+1, len(chunks))
    return term_counter, term_contexts

def extract_text_from_file(path):
    ext = os.path.splitext(path)[-1].lower()
    if ext == ".txt":
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    elif ext == ".docx":
        doc = Document(path)
        return "\n\n".join([para.text for para in doc.paragraphs if para.text.strip()])
    elif ext == ".xliff":
        tree = ET.parse(path)
        root = tree.getroot()
        ns = {"ns": root.tag.split('}')[0].strip('{')}
        sources = root.findall(".//ns:source", ns)
        return "\n\n".join([elem.text for elem in sources if elem.text])
    elif ext == ".html":
        with open(path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "html.parser")
            for tag in soup(["script", "style"]):
                tag.decompose()
            return soup.get_text(separator="\n")
    else:
        raise ValueError("Unsupported file format.")

# --- Tkinter GUI App ---

class TermExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Term Extractor")
        self.root.geometry("900x650")  # increased height for nav buttons

        self.nlp = spacy.load("en_core_web_sm")
        self.nlp.max_length = 2_000_000

        self.term_data = []  # list of dicts: {term, freq, contexts[], selected}
        self.term_data_sorted_asc = True

        # For context navigation:
        self.current_term = None
        self.current_context_index = 0

        self.create_widgets()

    def create_widgets(self):
        frm_top = tk.Frame(self.root)
        frm_top.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(frm_top, text="Input file:").pack(side=tk.LEFT)
        self.file_entry = tk.Entry(frm_top, width=60)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(frm_top, text="Browse", command=self.browse_file).pack(side=tk.LEFT)

        tk.Label(frm_top, text="Min frequency:").pack(side=tk.LEFT, padx=(20,5))
        self.min_freq_entry = tk.Entry(frm_top, width=5)
        self.min_freq_entry.insert(0, "1")
        self.min_freq_entry.pack(side=tk.LEFT)

        tk.Button(frm_top, text="Extract terms", command=self.extract_terms_thread).pack(side=tk.LEFT, padx=10)

        self.progress_var = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progressbar.pack(fill=tk.X, padx=5, pady=5)

        # Term counter label
        self.term_count_label = tk.Label(self.root, text="(0 unique terms extracted)")
        self.term_count_label.pack()

        # Treeview for terms
        columns = ("Selected", "Term", "Frequency", "Context")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
        self.tree.column("Selected", width=60, anchor="center")
        self.tree.column("Term", width=250)
        self.tree.column("Frequency", width=80, anchor="center")
        self.tree.column("Context", width=450)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tree.bind("<<TreeviewSelect>>", self.on_term_selected)
        self.tree.bind("<Button-1>", self.handle_checkbox_click)

        # Context display with highlighted term
        frm_context = tk.LabelFrame(self.root, text="Term Context")
        frm_context.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)

        self.context_text = tk.Text(frm_context, height=6, wrap=tk.WORD, bg="lightyellow")
        self.context_text.pack(fill=tk.BOTH, expand=True)

        # Navigation buttons for contexts
        nav_frame = tk.Frame(self.root)
        nav_frame.pack(pady=(0,10))

        self.prev_btn = tk.Button(nav_frame, text="← Previous Context", command=self.show_prev_context)
        self.prev_btn.pack(side=tk.LEFT, padx=10)

        self.next_btn = tk.Button(nav_frame, text="Next Context →", command=self.show_next_context)
        self.next_btn.pack(side=tk.LEFT, padx=10)

        # Export button
        btn_export = tk.Button(self.root, text="Export Selected to CSV", command=self.export_selected)
        btn_export.pack(pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Supported files", "*.txt *.docx *.xliff *.html"),
                       ("Text files", "*.txt"),
                       ("Word files", "*.docx"),
                       ("XLIFF files", "*.xliff"),
                       ("HTML files", "*.html")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def extract_terms_thread(self):
        # Run extraction in a thread to avoid freezing UI
        threading.Thread(target=self.extract_terms).start()

    def extract_terms(self):
        file_path = self.file_entry.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid input file.")
            return
        try:
            min_freq = int(self.min_freq_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Min frequency must be an integer.")
            return

        try:
            text = extract_text_from_file(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            return

        chunks = chunk_text(text)

        self.term_data.clear()
        self.progress_var.set(0)
        total_chunks = len(chunks)

        def progress_callback(completed, total):
            self.progress_var.set((completed/total)*100)

        counts, contexts = extract_terms_with_context(chunks, self.nlp, progress_callback)

        # Build term_data list with all terms above min_freq
        for term, freq in counts.items():
            if freq >= min_freq:
                self.term_data.append({
                    "term": term,
                    "freq": freq,
                    "contexts": contexts.get(term, []),  # now a list of contexts
                    "selected": True
                })

        self.term_data.sort(key=lambda x: x["freq"], reverse=not self.term_data_sorted_asc)

        self.update_treeview()
        self.progress_var.set(0)

        # Update term count label
        self.term_count_label.config(text=f"({len(self.term_data)} unique terms extracted)")

        # Reset navigation state
        self.current_term = None
        self.current_context_index = 0
        self.context_text.config(state=tk.NORMAL)
        self.context_text.delete("1.0", tk.END)
        self.context_text.config(state=tk.DISABLED)

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for idx, item in enumerate(self.term_data):
            checked = "✔" if item["selected"] else ""
            freq = item["freq"]
            # Show first context preview only
            first_context = item["contexts"][0] if item["contexts"] else ""
            context_preview = first_context[:80] + ("..." if len(first_context) > 80 else "")
            self.tree.insert("", "end", iid=idx, values=(checked, item["term"], freq, context_preview))

    def sort_by_column(self, col):
        if col == "Frequency":
            self.term_data_sorted_asc = not self.term_data_sorted_asc
            self.term_data.sort(key=lambda x: x["freq"], reverse=not self.term_data_sorted_asc)
            self.update_treeview()

    def on_term_selected(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        item = self.term_data[idx]
        term = item["term"]
        self.current_term = term
        self.current_context_index = 0
        contexts = item["contexts"]
        if contexts:
            self.show_context(term, contexts[self.current_context_index])
        else:
            self.context_text.config(state=tk.NORMAL)
            self.context_text.delete("1.0", tk.END)
            self.context_text.config(state=tk.DISABLED)

    def show_context(self, term, context):
        self.context_text.config(state=tk.NORMAL)
        self.context_text.delete("1.0", tk.END)

        # Insert context text with term highlighted (case insensitive)
        pattern = re.compile(re.escape(term), re.IGNORECASE)
        last_end = 0
        for m in pattern.finditer(context):
            self.context_text.insert(tk.END, context[last_end:m.start()])
            self.context_text.insert(tk.END, context[m.start():m.end()], "highlight")
            last_end = m.end()
        self.context_text.insert(tk.END, context[last_end:])
        self.context_text.tag_config("highlight", background="yellow")
        self.context_text.config(state=tk.DISABLED)

    def show_prev_context(self):
        if not self.current_term:
            return
        idx = next((i for i, d in enumerate(self.term_data) if d["term"] == self.current_term), None)
        if idx is None:
            return
        contexts = self.term_data[idx]["contexts"]
        if not contexts:
            return
        self.current_context_index = (self.current_context_index - 1) % len(contexts)
        self.show_context(self.current_term, contexts[self.current_context_index])

    def show_next_context(self):
        if not self.current_term:
            return
        idx = next((i for i, d in enumerate(self.term_data) if d["term"] == self.current_term), None)
        if idx is None:
            return
        contexts = self.term_data[idx]["contexts"]
        if not contexts:
            return
        self.current_context_index = (self.current_context_index + 1) % len(contexts)
        self.show_context(self.current_term, contexts[self.current_context_index])

    def handle_checkbox_click(self, event):
        # Detect clicks on the first column (checkboxes) to toggle selected state
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":  # First column is "Selected"
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        idx = int(row_id)
        self.term_data[idx]["selected"] = not self.term_data[idx]["selected"]
        self.update_treeview()
        # Keep selection on clicked row
        self.tree.selection_set(row_id)

    def export_selected(self):
        selected_terms = [t for t in self.term_data if t["selected"]]
        if not selected_terms:
            messagebox.showwarning("Export", "No terms selected for export.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save selected terms as CSV")
        if not file_path:
            return

        # Prepare DataFrame
        rows = []
        for item in selected_terms:
            rows.append({
                "Source term": item["term"],
                "Target term": "",
                "Frequency": item["freq"],
                "Context": item["contexts"][0] if item["contexts"] else "",
                "Notes": "Auto-extracted"
            })

        df = pd.DataFrame(rows)
        try:
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Export", f"Exported {len(rows)} terms to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TermExtractorApp(root)
    root.mainloop()

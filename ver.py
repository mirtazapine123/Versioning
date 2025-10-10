import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import sqlite3
from datetime import datetime
import os
import base64
from PIL import Image, ImageGrab, ImageTk
import io
from difflib import SequenceMatcher
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

class MachineTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema Tracciamento Modifiche Macchine - Versione Avanzata")
        self.root.geometry("1400x800")
        
        # Lista immagini per l'intervento corrente
        self.current_images = []
        self.preview_images = []
        
        # Inizializza database
        self.init_database()
        
        # Crea interfaccia
        self.create_widgets()
        
        # Carica dati iniziali
        self.load_all_records()
    
    def init_database(self):
        """Inizializza il database SQLite"""
        self.conn = sqlite3.connect('macchine_tracker.db')
        self.cursor = self.conn.cursor()
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS interventi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_ora TEXT NOT NULL,
                macchina TEXT NOT NULL,
                operatore TEXT NOT NULL,
                categoria TEXT NOT NULL,
                problema TEXT NOT NULL,
                soluzione TEXT NOT NULL
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS immagini (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                intervento_id INTEGER NOT NULL,
                nome_file TEXT NOT NULL,
                immagine BLOB NOT NULL,
                FOREIGN KEY (intervento_id) REFERENCES interventi(id) ON DELETE CASCADE
            )
        ''')
        
        self.conn.commit()
    
    def create_widgets(self):
        """Crea l'interfaccia grafica con tabs"""
        
        # Notebook per le tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Inserimento Interventi
        self.tab_insert = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_insert, text='Nuovo Intervento')
        self.create_insert_tab()
        
        # Tab 2: Ricerca e Storico
        self.tab_search = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_search, text='Ricerca e Storico')
        self.create_search_tab()
        
        # Tab 3: Assistente IA
        self.tab_ai = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_ai, text='Assistente IA')
        self.create_ai_tab()
        
        # Tab 4: Statistiche
        self.tab_stats = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stats, text='Statistiche')
        self.create_stats_tab()
    
    def create_insert_tab(self):
        """Crea la tab per inserimento dati"""
        main_frame = ttk.Frame(self.tab_insert, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.tab_insert.columnconfigure(0, weight=1)
        self.tab_insert.rowconfigure(0, weight=1)
        
        # Frame sinistro: campi input
        left_frame = ttk.LabelFrame(main_frame, text="Dati Intervento", padding="10")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        main_frame.columnconfigure(0, weight=2)
        main_frame.rowconfigure(0, weight=1)
        
        # Campo Macchina
        ttk.Label(left_frame, text="Macchina:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.macchina_entry = ttk.Entry(left_frame, width=50)
        self.macchina_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Campo Operatore
        ttk.Label(left_frame, text="Operatore:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.operatore_entry = ttk.Entry(left_frame, width=50)
        self.operatore_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Campo Categoria
        ttk.Label(left_frame, text="Categoria:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.categoria_combo = ttk.Combobox(left_frame, width=47, state="readonly")
        self.categoria_combo['values'] = ('Software', 'Elettrico', 'Pneumatico', 'Meccanico', 'Manutenzione', 'Altro')
        self.categoria_combo.current(0)
        self.categoria_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Campo Problema
        ttk.Label(left_frame, text="Problema:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.problema_text = scrolledtext.ScrolledText(left_frame, width=50, height=8, wrap=tk.WORD)
        self.problema_text.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Campo Soluzione
        ttk.Label(left_frame, text="Soluzione:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.soluzione_text = scrolledtext.ScrolledText(left_frame, width=50, height=8, wrap=tk.WORD)
        self.soluzione_text.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Pulsanti salvataggio
        button_frame = ttk.Frame(left_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="💾 Salva Intervento", command=self.save_record).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🗑️ Pulisci Campi", command=self.clear_fields).pack(side=tk.LEFT, padx=5)
        
        left_frame.columnconfigure(1, weight=1)
        
        # Frame destro: gestione immagini
        right_frame = ttk.LabelFrame(main_frame, text="Allegati Immagini", padding="10")
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        main_frame.columnconfigure(1, weight=1)
        
        # Pulsanti per immagini
        img_buttons = ttk.Frame(right_frame)
        img_buttons.pack(fill='x', pady=(0, 10))
        
        ttk.Button(img_buttons, text="📷 Screenshot", command=self.take_screenshot).pack(side=tk.LEFT, padx=5)
        ttk.Button(img_buttons, text="📁 Carica Immagine", command=self.load_image).pack(side=tk.LEFT, padx=5)
        ttk.Button(img_buttons, text="❌ Rimuovi Selezionata", command=self.remove_image).pack(side=tk.LEFT, padx=5)
        
        # Frame per preview immagini
        preview_frame = ttk.Frame(right_frame)
        preview_frame.pack(fill='both', expand=True)
        
        # Canvas con scrollbar per le preview
        canvas = tk.Canvas(preview_frame, height=500)
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=canvas.yview)
        self.preview_container = ttk.Frame(canvas)
        
        self.preview_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.preview_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def create_search_tab(self):
        """Crea la tab per ricerca e visualizzazione"""
        main_frame = ttk.Frame(self.tab_search, padding="10")
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Barra ricerca
        search_frame = ttk.LabelFrame(main_frame, text="Ricerca", padding="10")
        search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        
        ttk.Label(search_frame, text="Cerca:").grid(row=0, column=0, padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        self.search_entry.bind('<KeyRelease>', lambda e: self.search_records())
        
        ttk.Button(search_frame, text="🔍 Cerca", command=self.search_records).grid(row=0, column=2, padx=5)
        ttk.Button(search_frame, text="📋 Mostra Tutti", command=self.load_all_records).grid(row=0, column=3, padx=5)
        ttk.Button(search_frame, text="📊 Export Excel", command=self.export_to_excel).grid(row=0, column=4, padx=5)
        
        # Frame risultati
        results_frame = ttk.Frame(main_frame)
        results_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Treeview
        tree_scroll = ttk.Scrollbar(results_frame)
        tree_scroll.pack(side='right', fill='y')
        
        self.tree = ttk.Treeview(results_frame, yscrollcommand=tree_scroll.set, selectmode='browse')
        self.tree.pack(side='left', fill='both', expand=True)
        tree_scroll.config(command=self.tree.yview)
        
        self.tree['columns'] = ('Data', 'Macchina', 'Operatore', 'Categoria', 'Problema')
        self.tree.column('#0', width=0, stretch=tk.NO)
        self.tree.column('Data', anchor=tk.W, width=130)
        self.tree.column('Macchina', anchor=tk.W, width=120)
        self.tree.column('Operatore', anchor=tk.W, width=100)
        self.tree.column('Categoria', anchor=tk.W, width=100)
        self.tree.column('Problema', anchor=tk.W, width=400)
        
        self.tree.heading('Data', text='Data/Ora')
        self.tree.heading('Macchina', text='Macchina')
        self.tree.heading('Operatore', text='Operatore')
        self.tree.heading('Categoria', text='Categoria')
        self.tree.heading('Problema', text='Problema')
        
        self.tree.bind('<Double-1>', self.show_details)
        
        # Dettagli
        details_frame = ttk.LabelFrame(main_frame, text="Dettagli Intervento", padding="10")
        details_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.details_text = scrolledtext.ScrolledText(details_frame, height=10, wrap=tk.WORD, state='disabled')
        self.details_text.pack(fill='both', expand=True)
        
        btn_frame = ttk.Frame(details_frame)
        btn_frame.pack(pady=(10, 0))
        
        ttk.Button(btn_frame, text="🖼️ Visualizza Immagini", command=self.view_images).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ Elimina Intervento", command=self.delete_record).pack(side=tk.LEFT, padx=5)
    
    def create_ai_tab(self):
        """Crea la tab per l'assistente IA"""
        main_frame = ttk.Frame(self.tab_ai, padding="10")
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Intestazione
        header = ttk.Label(main_frame, text="🤖 Assistente IA - Trova Soluzioni Simili", 
                          font=('Arial', 14, 'bold'))
        header.grid(row=0, column=0, pady=(0, 10))
        
        # Input domanda
        input_frame = ttk.LabelFrame(main_frame, text="Descrivi il Problema", padding="10")
        input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        
        self.ai_question = scrolledtext.ScrolledText(input_frame, height=4, wrap=tk.WORD)
        self.ai_question.pack(fill='x', pady=(0, 10))
        
        btn_frame = ttk.Frame(input_frame)
        btn_frame.pack()
        
        ttk.Button(btn_frame, text="🔍 Cerca Soluzioni Simili", command=self.ai_find_solutions).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ Pulisci", command=lambda: self.ai_question.delete('1.0', tk.END)).pack(side=tk.LEFT, padx=5)
        
        # Risultati
        results_frame = ttk.LabelFrame(main_frame, text="Soluzioni Trovate", padding="10")
        results_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        self.ai_results = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, state='disabled')
        self.ai_results.pack(fill='both', expand=True)
    
    def create_stats_tab(self):
        """Crea la tab per le statistiche"""
        main_frame = ttk.Frame(self.tab_stats, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Pulsante aggiorna
        ttk.Button(main_frame, text="🔄 Aggiorna Statistiche", command=self.update_statistics).pack(pady=10)
        
        # Frame per i grafici
        self.stats_container = ttk.Frame(main_frame)
        self.stats_container.pack(fill='both', expand=True)
        
        self.update_statistics()
    
    def take_screenshot(self):
        """Cattura uno screenshot"""
        self.root.withdraw()
        self.root.after(500, self._capture_screen)
    
    def _capture_screen(self):
        """Cattura effettivamente lo schermo dopo un delay"""
        try:
            screenshot = ImageGrab.grab()
            
            # Ridimensiona per preview
            screenshot.thumbnail((300, 300))
            
            # Salva in memoria
            img_byte_arr = io.BytesIO()
            screenshot.save(img_byte_arr, format='PNG')
            img_data = img_byte_arr.getvalue()
            
            self.current_images.append({
                'name': f'screenshot_{len(self.current_images)+1}.png',
                'data': img_data
            })
            
            self.update_image_preview()
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante lo screenshot: {e}")
        finally:
            self.root.deiconify()
    
    def load_image(self):
        """Carica un'immagine da file"""
        file_path = filedialog.askopenfilename(
            title="Seleziona Immagine",
            filetypes=[("Immagini", "*.png *.jpg *.jpeg *.gif *.bmp"), ("Tutti i file", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'rb') as f:
                    img_data = f.read()
                
                file_name = os.path.basename(file_path)
                
                self.current_images.append({
                    'name': file_name,
                    'data': img_data
                })
                
                self.update_image_preview()
                
            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel caricamento: {e}")
    
    def remove_image(self):
        """Rimuove l'immagine selezionata"""
        if not self.current_images:
            messagebox.showwarning("Attenzione", "Nessuna immagine da rimuovere!")
            return
        
        # Finestra per selezionare quale rimuovere
        if len(self.current_images) == 1:
            self.current_images.pop(0)
        else:
            # Dialog semplice
            idx = messagebox.askquestion("Rimuovi", 
                f"Rimuovere l'ultima immagine ({self.current_images[-1]['name']})?")
            if idx == 'yes':
                self.current_images.pop(-1)
        
        self.update_image_preview()
    
    def update_image_preview(self):
        """Aggiorna la preview delle immagini"""
        # Pulisci preview esistenti
        for widget in self.preview_container.winfo_children():
            widget.destroy()
        
        self.preview_images = []
        
        for idx, img_data in enumerate(self.current_images):
            frame = ttk.LabelFrame(self.preview_container, text=img_data['name'], padding="5")
            frame.pack(fill='x', pady=5)
            
            try:
                image = Image.open(io.BytesIO(img_data['data']))
                image.thumbnail((250, 250))
                
                photo = ImageTk.PhotoImage(image)
                self.preview_images.append(photo)
                
                label = ttk.Label(frame, image=photo)
                label.pack()
                
            except Exception as e:
                ttk.Label(frame, text=f"Errore visualizzazione: {e}").pack()
    
    def save_record(self):
        """Salva un nuovo intervento nel database"""
        macchina = self.macchina_entry.get().strip()
        operatore = self.operatore_entry.get().strip()
        categoria = self.categoria_combo.get()
        problema = self.problema_text.get('1.0', tk.END).strip()
        soluzione = self.soluzione_text.get('1.0', tk.END).strip()
        
        if not macchina or not operatore or not problema or not soluzione:
            messagebox.showwarning("Campi Mancanti", "Compila tutti i campi obbligatori!")
            return
        
        data_ora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            # Inserisci intervento
            self.cursor.execute('''
                INSERT INTO interventi (data_ora, macchina, operatore, categoria, problema, soluzione)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (data_ora, macchina, operatore, categoria, problema, soluzione))
            
            intervento_id = self.cursor.lastrowid
            
            # Salva immagini
            for img in self.current_images:
                self.cursor.execute('''
                    INSERT INTO immagini (intervento_id, nome_file, immagine)
                    VALUES (?, ?, ?)
                ''', (intervento_id, img['name'], img['data']))
            
            self.conn.commit()
            messagebox.showinfo("Successo", f"Intervento salvato con {len(self.current_images)} immagini!")
            
            self.clear_fields()
            self.load_all_records()
            
        except sqlite3.Error as e:
            messagebox.showerror("Errore Database", f"Errore: {e}")
    
    def clear_fields(self):
        """Pulisce tutti i campi"""
        self.macchina_entry.delete(0, tk.END)
        self.operatore_entry.delete(0, tk.END)
        self.categoria_combo.current(0)
        self.problema_text.delete('1.0', tk.END)
        self.soluzione_text.delete('1.0', tk.END)
        self.current_images = []
        self.update_image_preview()
    
    def load_all_records(self):
        """Carica tutti i record"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.cursor.execute('''
            SELECT id, data_ora, macchina, operatore, categoria, problema 
            FROM interventi 
            ORDER BY data_ora DESC
        ''')
        
        for row in self.cursor.fetchall():
            problema_short = row[5][:80] + '...' if len(row[5]) > 80 else row[5]
            self.tree.insert('', tk.END, iid=row[0], values=(row[1], row[2], row[3], row[4], problema_short))
    
    def search_records(self):
        """Cerca nei record"""
        search_term = self.search_entry.get().strip().lower()
        
        if not search_term:
            self.load_all_records()
            return
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.cursor.execute('''
            SELECT id, data_ora, macchina, operatore, categoria, problema 
            FROM interventi 
            WHERE LOWER(problema) LIKE ? OR LOWER(soluzione) LIKE ? OR LOWER(macchina) LIKE ?
            ORDER BY data_ora DESC
        ''', (f'%{search_term}%', f'%{search_term}%', f'%{search_term}%'))
        
        results = self.cursor.fetchall()
        
        if not results:
            messagebox.showinfo("Ricerca", "Nessun risultato trovato.")
            return
        
        for row in results:
            problema_short = row[5][:80] + '...' if len(row[5]) > 80 else row[5]
            self.tree.insert('', tk.END, iid=row[0], values=(row[1], row[2], row[3], row[4], problema_short))
    
    def show_details(self, event):
        """Mostra dettagli intervento"""
        selection = self.tree.selection()
        if not selection:
            return
        
        record_id = selection[0]
        
        self.cursor.execute('''
            SELECT data_ora, macchina, operatore, categoria, problema, soluzione 
            FROM interventi 
            WHERE id = ?
        ''', (record_id,))
        
        record = self.cursor.fetchone()
        
        if record:
            # Conta immagini
            self.cursor.execute('SELECT COUNT(*) FROM immagini WHERE intervento_id = ?', (record_id,))
            num_images = self.cursor.fetchone()[0]
            
            details = f"""DATA/ORA: {record[0]}
MACCHINA: {record[1]}
OPERATORE: {record[2]}
CATEGORIA: {record[3]}
IMMAGINI ALLEGATE: {num_images}

PROBLEMA:
{record[4]}

SOLUZIONE:
{record[5]}"""
            
            self.details_text.config(state='normal')
            self.details_text.delete('1.0', tk.END)
            self.details_text.insert('1.0', details)
            self.details_text.config(state='disabled')
    
    def view_images(self):
        """Visualizza le immagini dell'intervento selezionato"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Selezione", "Seleziona un intervento!")
            return
        
        record_id = selection[0]
        
        self.cursor.execute('SELECT nome_file, immagine FROM immagini WHERE intervento_id = ?', (record_id,))
        images = self.cursor.fetchall()
        
        if not images:
            messagebox.showinfo("Immagini", "Nessuna immagine allegata a questo intervento.")
            return
        
        # Finestra per visualizzare immagini
        img_window = tk.Toplevel(self.root)
        img_window.title("Immagini Allegate")
        img_window.geometry("800x600")
        
        canvas = tk.Canvas(img_window)
        scrollbar = ttk.Scrollbar(img_window, orient="vertical", command=canvas.yview)
        container = ttk.Frame(canvas)
        
        container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        temp_photos = []
        
        for idx, (nome, img_data) in enumerate(images):
            frame = ttk.LabelFrame(container, text=nome, padding="10")
            frame.pack(fill='x', pady=10, padx=10)
            
            try:
                image = Image.open(io.BytesIO(img_data))
                image.thumbnail((700, 700))
                
                photo = ImageTk.PhotoImage(image)
                temp_photos.append(photo)
                
                label = ttk.Label(frame, image=photo)
                label.pack()
                
            except Exception as e:
                ttk.Label(frame, text=f"Errore: {e}").pack()
        
        # Mantieni riferimento
        img_window.photos = temp_photos
    
    def delete_record(self):
        """Elimina intervento"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Selezione", "Seleziona un intervento!")
            return
        
        if messagebox.askyesno("Conferma", "Eliminare questo intervento e le sue immagini?"):
            record_id = selection[0]
            
            try:
                self.cursor.execute('DELETE FROM immagini WHERE intervento_id = ?', (record_id,))
                self.cursor.execute('DELETE FROM interventi WHERE id = ?', (record_id,))
                self.conn.commit()
                
                messagebox.showinfo("Successo", "Intervento eliminato!")
                self.load_all_records()
                
                self.details_text.config(state='normal')
                self.details_text.delete('1.0', tk.END)
                self.details_text.config(state='disabled')
                
            except sqlite3.Error as e:
                messagebox.showerror("Errore", f"Errore: {e}")
    
    def ai_find_solutions(self):
        """IA: trova soluzioni simili"""
        question = self.ai_question.get('1.0', tk.END).strip()
        
        if not question:
            messagebox.showwarning("Attenzione", "Inserisci una descrizione del problema!")
            return
        
        # Carica tutti gli interventi
        self.cursor.execute('SELECT id, problema, soluzione, macchina, categoria, data_ora FROM interventi')
        all_records = self.cursor.fetchall()
        
        if not all_records:
            messagebox.showinfo("IA", "Nessun intervento nel database.")
            return
        
        # Calcola similarità
        similarities = []
        for record in all_records:
            similarity = self.calculate_similarity(question.lower(), record[1].lower())
            if similarity > 0.3:  # Soglia minima
                similarities.append((similarity, record))
        
        # Ordina per similarità
        similarities.sort(reverse=True, key=lambda x: x[0])
        
        # Mostra risultati
        self.ai_results.config(state='normal')
        self.ai_results.delete('1.0', tk.END)
        
        if not similarities:
            self.ai_results.insert('1.0', "❌ Nessuna soluzione simile trovata nel database.\n\n")
            self.ai_results.insert(tk.END, "Suggerimenti:\n")
            self.ai_results.insert(tk.END, "- Prova a descrivere il problema in modo diverso\n")
            self.ai_results.insert(tk.END, "- Usa parole chiave più generiche\n")
            self.ai_results.insert(tk.END, "- Aggiungi più interventi al database per migliorare i risultati")
        else:
            self.ai_results.insert('1.0', f"✅ Trovate {len(similarities[:5])} soluzioni simili:\n\n")
            self.ai_results.insert(tk.END, "="*80 + "\n\n")
            
            for idx, (similarity, record) in enumerate(similarities[:5]):
                record_id, problema, soluzione, macchina, categoria, data_ora = record
                percentage = int(similarity * 100)
                
                self.ai_results.insert(tk.END, f"🔍 RISULTATO #{idx+1} - Similarità: {percentage}%\n")
                self.ai_results.insert(tk.END, f"{'─'*80}\n")
                self.ai_results.insert(tk.END, f"📅 Data: {data_ora}\n")
                self.ai_results.insert(tk.END, f"🔧 Macchina: {macchina}\n")
                self.ai_results.insert(tk.END, f"📂 Categoria: {categoria}\n\n")
                self.ai_results.insert(tk.END, f"❓ PROBLEMA:\n{problema}\n\n")
                self.ai_results.insert(tk.END, f"✅ SOLUZIONE:\n{soluzione}\n\n")
                self.ai_results.insert(tk.END, "="*80 + "\n\n")
        
        self.ai_results.config(state='disabled')
    
    def calculate_similarity(self, text1, text2):
        """Calcola similarità tra due testi"""
        return SequenceMatcher(None, text1, text2).ratio()
    
    def update_statistics(self):
        """Aggiorna i grafici statistici"""
        # Pulisci container
        for widget in self.stats_container.winfo_children():
            widget.destroy()
        
        # Carica dati
        self.cursor.execute('SELECT COUNT(*) FROM interventi')
        total = self.cursor.fetchone()[0]
        
        if total == 0:
            ttk.Label(self.stats_container, text="Nessun dato disponibile", 
                     font=('Arial', 14)).pack(pady=50)
            return
        
        # Frame per info generali
        info_frame = ttk.LabelFrame(self.stats_container, text="Informazioni Generali", padding="15")
        info_frame.pack(fill='x', padx=10, pady=10)
        
        self.cursor.execute('SELECT COUNT(*) FROM immagini')
        total_images = self.cursor.fetchone()[0]
        
        self.cursor.execute('SELECT COUNT(DISTINCT macchina) FROM interventi')
        unique_machines = self.cursor.fetchone()[0]
        
        info_text = f"""
        📊 Totale Interventi: {total}
        🖼️  Totale Immagini: {total_images}
        🔧 Macchine Diverse: {unique_machines}
        📈 Media Immagini/Intervento: {total_images/total if total > 0 else 0:.1f}
        """
        
        ttk.Label(info_frame, text=info_text, font=('Arial', 11)).pack()
        
        # Frame per grafici
        charts_frame = ttk.Frame(self.stats_container)
        charts_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Grafico 1: Interventi per categoria
        self.cursor.execute('''
            SELECT categoria, COUNT(*) 
            FROM interventi 
            GROUP BY categoria 
            ORDER BY COUNT(*) DESC
        ''')
        cat_data = self.cursor.fetchall()
        
        if cat_data:
            fig1, ax1 = plt.subplots(figsize=(6, 4))
            categories = [row[0] for row in cat_data]
            counts = [row[1] for row in cat_data]
            
            colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A', '#98D8C8', '#C7CEEA']
            ax1.bar(categories, counts, color=colors[:len(categories)])
            ax1.set_title('Interventi per Categoria', fontsize=14, fontweight='bold')
            ax1.set_xlabel('Categoria')
            ax1.set_ylabel('Numero Interventi')
            ax1.tick_params(axis='x', rotation=45)
            plt.tight_layout()
            
            canvas1 = FigureCanvasTkAgg(fig1, charts_frame)
            canvas1.draw()
            canvas1.get_tk_widget().pack(side=tk.LEFT, fill='both', expand=True, padx=5)
        
        # Grafico 2: Top 5 macchine con più interventi
        self.cursor.execute('''
            SELECT macchina, COUNT(*) as cnt
            FROM interventi 
            GROUP BY macchina 
            ORDER BY cnt DESC 
            LIMIT 5
        ''')
        machine_data = self.cursor.fetchall()
        
        if machine_data:
            fig2, ax2 = plt.subplots(figsize=(6, 4))
            machines = [row[0][:15] + '...' if len(row[0]) > 15 else row[0] for row in machine_data]
            counts = [row[1] for row in machine_data]
            
            ax2.barh(machines, counts, color='#FF6B6B')
            ax2.set_title('Top 5 Macchine - Interventi', fontsize=14, fontweight='bold')
            ax2.set_xlabel('Numero Interventi')
            ax2.invert_yaxis()
            plt.tight_layout()
            
            canvas2 = FigureCanvasTkAgg(fig2, charts_frame)
            canvas2.draw()
            canvas2.get_tk_widget().pack(side=tk.LEFT, fill='both', expand=True, padx=5)
        
        # Grafico 3: Interventi per mese
        self.cursor.execute('''
            SELECT strftime('%Y-%m', data_ora) as mese, COUNT(*) 
            FROM interventi 
            GROUP BY mese 
            ORDER BY mese DESC 
            LIMIT 12
        ''')
        month_data = self.cursor.fetchall()
        
        if month_data and len(month_data) > 1:
            fig3, ax3 = plt.subplots(figsize=(12, 4))
            months = [row[0] for row in reversed(month_data)]
            counts = [row[1] for row in reversed(month_data)]
            
            ax3.plot(months, counts, marker='o', linewidth=2, markersize=8, color='#4ECDC4')
            ax3.fill_between(range(len(months)), counts, alpha=0.3, color='#4ECDC4')
            ax3.set_title('Trend Interventi per Mese', fontsize=14, fontweight='bold')
            ax3.set_xlabel('Mese')
            ax3.set_ylabel('Numero Interventi')
            ax3.tick_params(axis='x', rotation=45)
            ax3.grid(True, alpha=0.3)
            plt.tight_layout()
            
            canvas3 = FigureCanvasTkAgg(fig3, self.stats_container)
            canvas3.draw()
            canvas3.get_tk_widget().pack(fill='both', expand=True, padx=10, pady=10)
    
    def export_to_excel(self):
        """Esporta tutti i dati in Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"export_interventi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not file_path:
            return
        
        try:
            # Carica tutti i dati
            self.cursor.execute('''
                SELECT id, data_ora, macchina, operatore, categoria, problema, soluzione 
                FROM interventi 
                ORDER BY data_ora DESC
            ''')
            records = self.cursor.fetchall()
            
            # Crea workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Interventi"
            
            # Stile header
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=12)
            
            # Headers
            headers = ['ID', 'Data/Ora', 'Macchina', 'Operatore', 'Categoria', 'Problema', 'Soluzione', 'N. Immagini']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col)
                cell.value = header
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Dati
            for row_idx, record in enumerate(records, 2):
                # Conta immagini
                self.cursor.execute('SELECT COUNT(*) FROM immagini WHERE intervento_id = ?', (record[0],))
                num_images = self.cursor.fetchone()[0]
                
                for col_idx, value in enumerate(record, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.value = value
                    cell.alignment = Alignment(vertical='top', wrap_text=True)
                
                # Aggiungi conteggio immagini
                ws.cell(row=row_idx, column=8).value = num_images
            
            # Imposta larghezza colonne
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 50
            ws.column_dimensions['G'].width = 50
            ws.column_dimensions['H'].width = 12
            
            # Foglio statistiche
            ws_stats = wb.create_sheet("Statistiche")
            
            # Statistiche per categoria
            ws_stats['A1'] = 'STATISTICHE PER CATEGORIA'
            ws_stats['A1'].font = Font(bold=True, size=14)
            ws_stats['A3'] = 'Categoria'
            ws_stats['B3'] = 'Conteggio'
            
            self.cursor.execute('SELECT categoria, COUNT(*) FROM interventi GROUP BY categoria ORDER BY COUNT(*) DESC')
            for idx, (cat, count) in enumerate(self.cursor.fetchall(), 4):
                ws_stats[f'A{idx}'] = cat
                ws_stats[f'B{idx}'] = count
            
            # Statistiche per macchina
            start_row = idx + 3
            ws_stats[f'A{start_row}'] = 'TOP 10 MACCHINE'
            ws_stats[f'A{start_row}'].font = Font(bold=True, size=14)
            ws_stats[f'A{start_row+2}'] = 'Macchina'
            ws_stats[f'B{start_row+2}'] = 'Interventi'
            
            self.cursor.execute('SELECT macchina, COUNT(*) FROM interventi GROUP BY macchina ORDER BY COUNT(*) DESC LIMIT 10')
            for idx, (machine, count) in enumerate(self.cursor.fetchall(), start_row+3):
                ws_stats[f'A{idx}'] = machine
                ws_stats[f'B{idx}'] = count
            
            ws_stats.column_dimensions['A'].width = 30
            ws_stats.column_dimensions['B'].width = 15
            
            # Salva
            wb.save(file_path)
            messagebox.showinfo("Successo", f"Dati esportati con successo!\n\n{len(records)} interventi salvati in:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Errore Export", f"Errore durante l'export: {e}")
    
    def __del__(self):
        """Chiude connessione database"""
        if hasattr(self, 'conn'):
            self.conn.close()

def main():
    root = tk.Tk()
    app = MachineTrackerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
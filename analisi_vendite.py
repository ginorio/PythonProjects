import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyodbc
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from datetime import datetime
from PIL import Image, ImageTk
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import io
import os

class SalesAnalysisApp:
    def __init__(self, master):
        self.master = master
        master.title("Analisi Vendite")
        master.geometry("1200x900")
        master.configure(bg='#f0f0f0')
        
        self.font = ("Segoe UI", 10)
        self.colors = {'I': '#4CAF50', 'E': '#2196F3'}
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TButton', font=('Segoe UI', 12), padding=10)
        self.style.configure('TButton', background='white', foreground='black')    #4CAF50
        self.style.map('TButton', background=[('active', '#45a049')], foreground=[('active', 'white')])
        self.style.configure('TLabel', font=self.font, background='#f0f0f0')
        self.style.configure('TEntry', font=self.font)
        self.style.configure('Treeview', font=('Segoe UI', 12))
        self.style.configure('Treeview.Heading', font=('Segoe UI', 12, 'bold'))
        
        self.main_frame = ttk.Frame(self.master, style='TFrame')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.max_y_value = 10000
        
        # Aggiunta della versione in alto a destra
        version_label = ttk.Label(self.main_frame, text="version 1.3 by Alberto Idrio", font=("Segoe UI", 8))
        version_label.pack(side=tk.TOP, anchor=tk.NE)
        
        self.create_date_inputs()
        self.create_buttons()
        self.create_pivot_table()
        self.create_detail_pivot_table()
        self.create_bar_chart()
        
        today = datetime.now()
        start_of_year = datetime(today.year, 1, 1)
        self.start_date.set_date(start_of_year)
        self.end_date.set_date(today)
        
    '''def create_date_inputs(self):
        date_frame = ttk.Frame(self.main_frame, style='TFrame')
        date_frame.pack(pady=10)
        
        ttk.Label(date_frame, text="Data inizio:", style='TLabel').grid(row=0, column=0, padx=5)
        self.start_date = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.start_date.grid(row=0, column=1, padx=5)
        
        ttk.Label(date_frame, text="Data fine:", style='TLabel').grid(row=0, column=2, padx=5)
        self.end_date = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.end_date.grid(row=0, column=3, padx=5)'''
    
    def create_date_inputs(self):
        # Creazione dello stile personalizzato per le etichette
        style = ttk.Style()
        style.configure('LargeLabel.TLabel', font=('Helvetica', 12))  # Puoi regolare la dimensione (12) come desideri

        date_frame = ttk.Frame(self.main_frame, style='TFrame')
        date_frame.pack(pady=10)
    
        ttk.Label(date_frame, text=" ", style='LargeLabel.TLabel').grid(row=0, column=2, padx=20)
        
        ttk.Label(date_frame, text="Data inizio:", style='LargeLabel.TLabel').grid(row=0, column=0, padx=5)
        self.start_date = DateEntry(date_frame, width=12, background='darkgreen', foreground='white', 
                                borderwidth=2, date_pattern='dd/mm/yyyy', 
                                font=('Helvetica', 12))  # Aumenta la dimensione del font
        self.start_date.grid(row=0, column=1, padx=5)
    
        ttk.Label(date_frame, text="Data fine:", style='LargeLabel.TLabel').grid(row=0, column=3, padx=5)
        self.end_date = DateEntry(date_frame, width=12, background='darkgreen', foreground='white', 
                          borderwidth=2, date_pattern='dd/mm/yyyy', 
                          font=('Helvetica', 12))  # Aumenta la dimensione del font
        self.end_date.grid(row=0, column=4, padx=5)
        
        '''ttk.Label(date_frame, text="Data fine:", style='LargeLabel.TLabel').grid(row=0, column=2, padx=5)
        self.end_date = DateEntry(date_frame, width=12, background='darkblue', foreground='white', 
                              borderwidth=2, date_pattern='dd/mm/yyyy', 
                              font=('Helvetica', 12))  # Aumenta la dimensione del font
        self.end_date.grid(row=0, column=3, padx=5)'''
        
    def create_buttons(self):
        button_frame = ttk.Frame(self.main_frame, style='TFrame')
        button_frame.pack(pady=10)
        
        update_icon = self.load_icon("update_icon.png")
        self.update_btn = ttk.Button(button_frame, text="Aggiorna", command=self.update_data, image=update_icon, compound=tk.LEFT)
        self.update_btn.image = update_icon
        self.update_btn.pack(side=tk.LEFT, padx=5)
        
        excel_icon = self.load_icon("excel_icon.png")
        self.excel_btn = ttk.Button(button_frame, text="Esporta in Excel", command=self.export_to_excel, image=excel_icon, compound=tk.LEFT)
        self.excel_btn.image = excel_icon
        self.excel_btn.pack(side=tk.LEFT, padx=5)
        
    def load_icon(self, filename):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "icons", filename)
        with Image.open(icon_path) as img:
            img = img.resize((50, 50), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        
    def create_pivot_table(self):
        self.tree = ttk.Treeview(self.main_frame, columns=('Listino', 'Totale'), show='headings', style='Treeview', height=3)
        self.tree.heading('Listino', text='Listino')
        self.tree.heading('Totale', text='Totale')
        self.tree.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
    def create_detail_pivot_table(self):
        # Frame per contenere la tabella e la barra di scorrimento
        detail_frame = ttk.Frame(self.main_frame)
        detail_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Creiamo la nuova tabella pivot di dettaglio
        self.detail_tree = ttk.Treeview(detail_frame, columns=('TabellaOrdine', 'NumeroOrdine', 'Listino', 'ImportoTotale'), show='headings', style='Treeview', height=13)
        self.detail_tree.heading('TabellaOrdine', text='Tabella Ordine')
        self.detail_tree.heading('NumeroOrdine', text='Numero Ordine')
        self.detail_tree.heading('Listino', text='Listino')
        self.detail_tree.heading('ImportoTotale', text='Importo Totale')
        
        # Configuriamo le colonne per adattarsi al contenuto
        self.detail_tree.column('TabellaOrdine', width=100)
        self.detail_tree.column('NumeroOrdine', width=100)
        self.detail_tree.column('Listino', width=100)
        self.detail_tree.column('ImportoTotale', width=100)

        # Aggiungiamo una barra di scorrimento
        scrollbar = ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=self.detail_tree.yview)
        self.detail_tree.configure(yscroll=scrollbar.set)

        # Posizionamento della tabella e della barra di scorrimento
        self.detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def create_bar_chart(self):
        self.figure = Figure(figsize=(6, 4), dpi=100, facecolor='#f0f0f0')
        self.plot = self.figure.add_subplot(1, 1, 1)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.main_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        
    '''def create_bar_chart(self):
        chart_frame = ttk.Frame(self.main_frame)
        chart_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=1, pady=10)

        # Grafico a barre
        self.bar_figure = Figure(figsize=(6, 4), dpi=100, facecolor='#f0f0f0')
        self.bar_plot = self.bar_figure.add_subplot(1, 1, 1)
        self.bar_canvas = FigureCanvasTkAgg(self.bar_figure, master=chart_frame)
        self.bar_canvas.draw()
        self.bar_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        # Grafico a torta
        self.pie_figure = Figure(figsize=(4, 4), dpi=100, facecolor='#f0f0f0')
        self.pie_plot = self.pie_figure.add_subplot(1, 1, 1)
        self.pie_canvas = FigureCanvasTkAgg(self.pie_figure, master=chart_frame)
        self.pie_canvas.draw()
        self.pie_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=1)'''
        
    def update_data(self):
        try:
            start_date = datetime.strptime(self.start_date.get(), "%d/%m/%Y")
            end_date = datetime.strptime(self.end_date.get(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Errore", "Formato data non valido. Usa DD/MM/YYYY.")
            return

        try:
            conn = pyodbc.connect('DSN=ARCHIVIASA')
            cursor = conn.cursor()

            query = """
            SELECT v.NumeroListino, SUM(vr.ImportoScontato) as TotaleVendite
            FROM Vebolt v
            JOIN Vebolr vr ON v.CodiceAziendaBolla = vr.CodiceAzienda
                AND v.TabellaBolla = vr.TabellaBolla
                AND v.NumeroBolla = vr.NumeroBolla
                AND v.TabellaOrdine = vr.TabellaOrdini
                AND v.NumeroOrdine = vr.NumeroOrdine
            WHERE v.CodiceAziendaBolla = '1'
                AND v.DataDocumento BETWEEN ? AND ?
                AND (vr.CodiceArticolo LIKE '42E6%'
                OR vr.CodiceArticolo LIKE '50E6%'
                OR vr.CodiceArticolo LIKE '52N6%'
                OR vr.CodiceArticolo LIKE '53N6%'
                OR vr.CodiceArticolo LIKE '54N6%'
                OR vr.CodiceArticolo LIKE '39U6%'
                OR vr.CodiceArticolo LIKE '49U6%')
            GROUP BY v.NumeroListino
            """

            cursor.execute(query, (start_date, end_date))
            results = cursor.fetchall()

            self.data = {row.NumeroListino: row.TotaleVendite for row in results}
            
            # Query per la tabella pivot di dettaglio
            detail_query = """
            SELECT v.TabellaOrdine, v.NumeroOrdine, v.NumeroListino, SUM(vr.ImportoScontato) as ImportoTotale
            FROM Vebolt v
            JOIN Vebolr vr ON v.CodiceAziendaBolla = vr.CodiceAzienda
                AND v.TabellaBolla = vr.TabellaBolla
                AND v.NumeroBolla = vr.NumeroBolla
                AND v.TabellaOrdine = vr.TabellaOrdini
                AND v.NumeroOrdine = vr.NumeroOrdine
            WHERE v.CodiceAziendaBolla = '1'
                AND v.DataDocumento BETWEEN ? AND ?
                AND (vr.CodiceArticolo LIKE '42E6%'
                OR vr.CodiceArticolo LIKE '50E6%'
                OR vr.CodiceArticolo LIKE '52N6%'
                OR vr.CodiceArticolo LIKE '53N6%'
                OR vr.CodiceArticolo LIKE '54N6%'
                OR vr.CodiceArticolo LIKE '39U6%'
                OR vr.CodiceArticolo LIKE '49U6%')
            GROUP BY v.TabellaOrdine, v.NumeroOrdine, v.NumeroListino
            ORDER BY v.TabellaOrdine, v.NumeroOrdine
            """

            cursor.execute(detail_query, (start_date, end_date))
            detail_results = cursor.fetchall()

            self.detail_data = detail_results

            self.update_pivot_table()
            self.update_detail_pivot_table()
            self.update_bar_chart()
            

        except pyodbc.Error as e:
            messagebox.showerror("Errore Database", f"Errore nella connessione al database: {str(e)}")
        finally:
            if 'conn' in locals():
                conn.close()

    #def update_pivot_table(self):
    #   for i in self.tree.get_children():
    #       self.tree.delete(i)
    #   for listino, totale in self.data.items():
    #       self.tree.insert('', 'end', values=(listino, f"{totale:,.2f}".replace(',', '.')))
    
    def update_pivot_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        listino_map = {'E': 'Estero', 'I': 'Italia'}
        for listino, totale in self.data.items():
            listino_display = listino_map.get(listino, listino)
            self.tree.insert('', 'end', values=(listino_display, f"{totale:,.2f} €".replace(',', '.')))
            
    def update_detail_pivot_table(self):
        for i in self.detail_tree.get_children():
            self.detail_tree.delete(i)
        listino_map = {'E': 'Estero', 'I': 'Italia'}
        for row in self.detail_data:
            listino_display = listino_map.get(row.NumeroListino, row.NumeroListino)
            self.detail_tree.insert('', 'end', values=(
                row.TabellaOrdine,
                row.NumeroOrdine,
                listino_display,
                f"{row.ImportoTotale:,.2f} €".replace(',', '.')
            ))
        
    def update_bar_chart(self):
        self.plot.clear()
        listini = list(self.data.keys())
        totali = list(self.data.values())
        colors = [self.colors.get(listino, 'gray') for listino in listini]
        
        listino_map = {'E': 'Estero', 'I': 'Italia'}
        listini = [listino_map.get(l, l) for l in listini]
    
        bars = self.plot.bar(listini, totali, color=colors)
        self.plot.set_ylabel('Totale vendite')
        self.plot.set_title('Vendite per listino')
        
        # Impostiamo il limite fisso per l'asse Y
        self.plot.set_ylim(0, self.max_y_value)
        
        # Aggiungiamo le etichette dei valori sopra ogni barra
        for bar in bars:
            height = bar.get_height()
            self.plot.text(bar.get_x() + bar.get_width()/2., height,
                           f'{height:,.0f} €',
                           ha='center', va='bottom', rotation=0)
        
        # Formattazione dell'asse Y per mostrare i valori in migliaia o milioni
        self.plot.yaxis.set_major_formatter(plt.FuncFormatter(self.format_yaxis))
        
        self.canvas.draw()

    def format_yaxis(self, value, _):
        if value >= 1e6:
            return f'{value/1e6:.1f}M'
        elif value >= 1e3:
            return f'{value/1e3:.0f}K'
        else:
            return f'{value:.0f}'
    
    def export_to_excel(self):
        if not hasattr(self, 'data'):
            messagebox.showwarning("Attenzione", "Non ci sono dati da esportare. Aggiorna prima i dati.")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.DataFrame(list(self.data.items()), columns=['Listino', 'Totale Vendite'])
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Esportazione Completata", f"I dati sono stati esportati in {file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesAnalysisApp(root)
    root.mainloop()
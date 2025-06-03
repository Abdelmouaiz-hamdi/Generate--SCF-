import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import re  # regex
from lxml import etree as ET  # assure-toi que lxml est bien installé



class SCFReplacementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔧 Comparateur et Remplacement XML SCF Nokia")
        self.df = None
        self.xml_content = ""
        self.replacements = []

        self.excel_path = tk.StringVar()
        self.old_site = tk.StringVar()
        self.new_site = tk.StringVar()
        self.all_sites = []

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="📄 Fichier Excel :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(self.root, textvariable=self.excel_path, width=60).grid(row=0, column=1, padx=5)
        tk.Button(self.root, text="Parcourir", command=self.load_excel).grid(row=0, column=2, padx=5)

        tk.Label(self.root, text="🏗️ Site ancien :").grid(row=1, column=0, sticky="w", padx=5)
        self.old_site_cb = ttk.Combobox(self.root, textvariable=self.old_site)
        self.old_site_cb.grid(row=1, column=1, padx=5, pady=5)
        self.old_site_cb.bind("<KeyRelease>", self.filter_old_site)

        tk.Label(self.root, text="🏗️ Site nouveau :").grid(row=2, column=0, sticky="w", padx=5)
        self.new_site_cb = ttk.Combobox(self.root, textvariable=self.new_site)
        self.new_site_cb.grid(row=2, column=1, padx=5, pady=5)
        self.new_site_cb.bind("<KeyRelease>", self.filter_new_site)

        tk.Button(self.root, text="📂 Charger XML", command=self.load_xml).grid(row=3, column=0, pady=5)
        tk.Button(self.root, text="🔍 Générer les remplacements", command=self.compare_sites).grid(row=3, column=1, pady=5)
        tk.Button(self.root, text="💾 Appliquer les changements", command=self.apply_replacements).grid(row=3, column=2, pady=5)
        tk.Button(self.root, text="🧹 Nettoyer XML (classes)", command=self.open_class_cleaner_window).grid(row=3,
                                                                                                           column=3,
                                                                                                           padx=5,
                                                                                                           pady=5)

        columns = ("N°", "Paramètre", "Ancien", "Nouveau", "Trouvé dans XML")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=25)
        for col in columns:
            width = 50 if col == "N°" else 150
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width)
        self.tree.grid(row=4, column=0, columnspan=3, padx=5, pady=10)

    def filter_old_site(self, event):
        value = self.old_site.get().lower()
        filtered = [s for s in self.all_sites if value in s.lower()]
        self.old_site_cb['values'] = filtered

    def filter_new_site(self, event):
        value = self.new_site.get().lower()
        filtered = [s for s in self.all_sites if value in s.lower()]
        self.new_site_cb['values'] = filtered

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        self.excel_path.set(file_path)

        try:
            self.df = pd.read_excel(file_path, header=2)
            self.df.columns = self.df.columns.str.strip().str.replace('\xa0', '', regex=False)

            possible_site_cols = [c for c in self.df.columns if "site" in c.lower()]
            if not possible_site_cols:
                raise ValueError("Aucune colonne de site détectée (contenant 'site') dans le fichier Excel.")

            site_col = possible_site_cols[0]
            self.df[site_col] = self.df[site_col].astype(str)
            self.site_column = site_col
            self.all_sites = self.df[site_col].dropna().tolist()
            self.old_site_cb['values'] = self.all_sites
            self.new_site_cb['values'] = self.all_sites
            messagebox.showinfo("Succès", "Fichier Excel chargé avec succès.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lecture Excel :\n{e}")

    def load_xml(self):
        path = filedialog.askopenfilename(filetypes=[("Fichier XML", "*.xml")])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                self.xml_content = f.read()
            messagebox.showinfo("Succès", "Fichier XML chargé.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lecture XML :\n{e}")

    def clean_val(self, v):
        try:
            if isinstance(v, float) and v.is_integer():
                return str(int(v))
            return str(v).strip()
        except:
            return str(v).strip()

    def replace_lnBtsId(self, xml_content, old_value, new_value):
        return xml_content.replace(old_value, new_value)

    def compare_sites(self):
        self.tree.delete(*self.tree.get_children())
        self.replacements.clear()

        if self.df is None or not self.xml_content:
            messagebox.showwarning("Manque", "Chargez Excel et XML d'abord.")
            return

        site_old = self.old_site.get()
        site_new = self.new_site.get()
        if not site_old or not site_new:
            messagebox.showwarning("Sélection", "Choisissez deux sites.")
            return

        try:
            row_old = self.df[self.df[self.site_column] == site_old].iloc[0]
            row_new = self.df[self.df[self.site_column] == site_new].iloc[0]
            count = 1

            for col in self.df.columns:
                val_old = self.clean_val(row_old[col])
                val_new = self.clean_val(row_new[col])

                if (
                    val_old != val_new and
                    val_old not in ["", "nan", "NaN", "None", "#N/A"] and
                    val_old.upper() != "N/A"
                ):
                    tag = col.replace("$", "")
                    found_in_xml = "❌"

                    if tag.lower() == "lnbtsid":
                        found_in_xml = "✅" if val_old in self.xml_content else "❌"
                        self.replacements.append(("lnbtsid", val_old, val_new))

                    elif tag.lower().startswith("wncel"):
                        old_pattern = f'WNCEL-{val_old}"'
                        new_pattern = f'WNCEL-{val_new}"'
                        if old_pattern in self.xml_content:
                            found_in_xml = "✅"
                            self.replacements.append((old_pattern, new_pattern))

                    else:
                        patterns = [
                            (f'<p name="{tag}">{val_old}</p>', f'<p name="{tag}">{val_new}</p>'),
                            (f'<p>{val_old}</p>', f'<p>{val_new}</p>'),
                            (f'>{val_old}</p>', f'>{val_new}</p>'),
                            (f'-{val_old}', f'-{val_new}'),
                            (f'>{val_old}', f'>{val_new}'),
                            (val_old, val_new)
                        ]

                        for xml_old, xml_new in patterns:
                            if xml_old in self.xml_content:
                                self.replacements.append((xml_old, xml_new))
                                found_in_xml = "✅"
                                break

                    self.tree.insert("", tk.END, values=(count, col, val_old, val_new, found_in_xml))
                    count += 1

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur génération remplacements :\n{e}")

    def safe_replace(self, xml_content, old_val, new_val):
        pattern = rf'(?<![\d\w]){re.escape(old_val)}(?![\d\w])'
        return re.sub(pattern, new_val, xml_content)

    def apply_replacements(self):
        if not self.replacements:
            messagebox.showwarning("Aucun remplacement", "Aucun remplacement détecté.")
            return

        modified_xml = self.xml_content
        for item in self.replacements:
            if isinstance(item, tuple) and len(item) == 3 and item[0].lower() == "lnbtsid":
                _, old_val, new_val = item
                modified_xml = self.replace_lnBtsId(modified_xml, old_val, new_val)
            else:
                old, new = item
                modified_xml = self.safe_replace(modified_xml, old, new)

        save_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("Fichiers XML", "*.xml")])
        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(modified_xml)
            messagebox.showinfo("Succès", f"Fichier XML modifié sauvegardé sous :\n{save_path}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur sauvegarde XML :\n{e}")

    def open_class_cleaner_window(self):
        def browse_input():
            filename = filedialog.askopenfilename(filetypes=[("Fichiers XML", "*.xml")])
            if filename:
                input_var.set(filename)

        def browse_output():
            filename = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("Fichiers XML", "*.xml")])
            if filename:
                output_var.set(filename)

        def remove_classes_from_xml():
            input_file = input_var.get()
            output_file = output_var.get()
            classes_raw = classes_var.get()
            classes_to_remove = [cls.strip() for cls in classes_raw.split(',') if cls.strip()]

            if not input_file or not output_file or not classes_to_remove:
                messagebox.showerror("Erreur", "Tous les champs doivent être remplis.")
                return

            try:
                tree = ET.parse(input_file)
                root = tree.getroot()
                namespace = {'ns': 'raml21.xsd'}

                for parent in root.findall('.//ns:cmData', namespace):
                    for mo in list(parent.findall('.//ns:managedObject', namespace)):
                        mo_class = mo.get('class')
                        if mo_class and any(to_remove in mo_class for to_remove in classes_to_remove):
                            parent.remove(mo)

                tree.write(output_file, encoding='UTF-8', xml_declaration=True, pretty_print=True)
                messagebox.showinfo("Succès", "Nettoyage terminé. Le nouveau fichier XML est prêt.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur pendant le traitement :\n{e}")

        # Nouvelle fenêtre
        window = tk.Toplevel(self.root)
        window.title("🧹 Nettoyage XML : Suppression de classes")

        input_var = tk.StringVar()
        output_var = tk.StringVar()
        classes_var = tk.StringVar()

        tk.Label(window, text="Fichier XML source :").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        tk.Entry(window, textvariable=input_var, width=50).grid(row=0, column=1, padx=5)
        tk.Button(window, text="Parcourir", command=browse_input).grid(row=0, column=2, padx=5)

        tk.Label(window, text="Fichier XML de sortie :").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        tk.Entry(window, textvariable=output_var, width=50).grid(row=1, column=1, padx=5)
        tk.Button(window, text="Parcourir", command=browse_output).grid(row=1, column=2, padx=5)

        tk.Label(window, text="Classes à supprimer (séparées par des virgules) :").grid(row=2, column=0, columnspan=3,
                                                                                        sticky="w", padx=5)
        tk.Entry(window, textvariable=classes_var, width=70).grid(row=3, column=0, columnspan=3, padx=5, pady=5)
        classes_var.set("LNADJU, LNREL, LNAJDW")  # valeur par défaut

        tk.Button(window, text="Supprimer les classes", command=remove_classes_from_xml, bg="red", fg="white").grid(
            row=4, column=1, pady=10)



if __name__ == "__main__":
    root = tk.Tk()
    app = SCFReplacementApp(root)
    root.mainloop()

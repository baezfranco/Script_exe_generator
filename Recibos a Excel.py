import os
import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog, Button, Label, messagebox


def extract_data_from_pdfs(folder_path):
    data = []  # Lista para la primera tabla
    recent_data = []  # Lista para la segunda tabla
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    # Extraer datos de la primera página
                    first_page = pdf.pages[0].extract_text()
                    if first_page:
                        periodo = first_page.split("PERIODO FACTURADO:")[1].split("\n")[0].strip()
                        base = int(first_page.split("kWh base")[1].split("\n")[0].strip())
                        intermedio = int(first_page.split("kWh intermedia")[1].split("\n")[0].strip())
                        punta = int(first_page.split("kWh punta")[1].split("\n")[0].strip())
                        total_consumo = base + intermedio + punta
                        kvarh = int(first_page.split("kVArh")[1].split("\n")[0].strip())
                        fp = float(first_page.split("Factor de potencia %")[1].split("\n")[0].strip())
                        kw_base = int(first_page.split("kW base")[1].split("\n")[0].strip())
                        kw_intermedio = int(first_page.split("kW intermedia")[1].split("\n")[0].strip())
                        kw_punta = int(first_page.split("kW punta")[1].split("\n")[0].strip())
                        kw_max = max(kw_base, kw_intermedio, kw_punta)
                        data.append({
                            "Periodo": periodo,
                            "Consumo Base (kWh)": base,
                            "Consumo Intermedio (kWh)": intermedio,
                            "Consumo Punta (kWh)": punta,
                            "Total Consumo (kWh)": total_consumo,
                            "kVArh": kvarh,
                            "Factor de Potencia": fp,
                            "Demanda Base (kW)": kw_base,
                            "Demanda Intermedio (kW)": kw_intermedio,
                            "Demanda Punta (kW)": kw_punta,
                            "Demanda Máxima (kW)": kw_max
                        })

                    # Extraer datos de la segunda página
                    if len(pdf.pages) > 1:
                        second_page = pdf.pages[1].extract_text()
                        if second_page:
                            rows = second_page.splitlines()
                            for row in rows[3:]:
                                cols = row.split()
                                if len(cols) >= 6:
                                    mes, demanda, consumo, fp, precio = cols[0], cols[1], cols[2], cols[3], cols[4]
                                    recent_data.append({
                                        "Mes": mes,
                                        "Demanda (kW)": demanda,
                                        "Consumo Total (kWh)": consumo,
                                        "Factor de Potencia": fp,
                                        "Precio Medio ($/kWh)": precio
                                    })
            except Exception as e:
                print(f"Error procesando {filename}: {e}")
    return data, recent_data


def main():
    # Configuración de la interfaz gráfica
    root = Tk()
    root.title("Extractor de Datos de Recibos GDMTH")
    root.geometry("400x200")

    def seleccionar_carpeta():
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta con PDFs")
        return carpeta

    def procesar():
        input_folder = seleccionar_carpeta()
        if not input_folder:
            messagebox.showerror("Error", "No seleccionaste una carpeta de entrada.")
            return

        output_folder = filedialog.askdirectory(title="Seleccionar carpeta para guardar resultados")
        if not output_folder:
            messagebox.showerror("Error", "No seleccionaste una carpeta de salida.")
            return

        data, recent_data = extract_data_from_pdfs(input_folder)
        if data:
            # Guardar primera tabla
            df = pd.DataFrame(data)
            df.to_excel(os.path.join(output_folder, "tabla_por_periodo.xlsx"), index=False)
        if recent_data:
            # Guardar segunda tabla
            df_recent = pd.DataFrame(recent_data)
            df_recent.to_excel(os.path.join(output_folder, "tabla_reciente.xlsx"), index=False)

        messagebox.showinfo("Éxito", "Los archivos se han generado correctamente.")

    # Botones e interfaz
    Button(root, text="Seleccionar carpeta y procesar", command=procesar, width=40).pack(pady=20)
    Button(root, text="Salir", command=root.quit, width=20).pack(pady=20)

    root.mainloop()


if __name__ == "__main__":
    main()
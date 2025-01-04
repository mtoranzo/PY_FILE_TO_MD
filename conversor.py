import os
import win32com.client as win32  # Solo necesario para convertir .doc a .docx en Windows
from markitdown import MarkItDown
import pandas as pd

def convert_doc_to_docx(doc_file):
    """
    Convierte un archivo .doc a .docx usando Microsoft Word (requiere Word instalado).
    """
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(doc_file)
    docx_file = doc_file + "x"
    doc.SaveAs(docx_file, FileFormat=16)  # 16 es el formato .docx
    doc.Close()
    word.Quit()
    return docx_file

def convert_xls_to_xlsx(xls_file):
    """
    Convierte un archivo .xls a .xlsx usando pandas.
    """
    xls = pd.ExcelFile(xls_file)
    xlsx_file = xls_file + "x"
    with pd.ExcelWriter(xlsx_file) as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return xlsx_file

def process_directory(input_dir, output_dir):
    """
    Procesa todos los archivos soportados en un directorio y los convierte a .md.
    """
    # Crear el directorio de salida si no existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Crear una instancia de MarkItDown
    markitdown = MarkItDown()

    # Recorrer todos los archivos en el directorio de entrada
    for filename in os.listdir(input_dir):
        input_file = os.path.join(input_dir, filename)
        
        # Determinar el tipo de archivo
        if filename.endswith(".doc"):
            # Convertir .doc a .docx
            docx_file = convert_doc_to_docx(input_file)
            input_file = docx_file  # Reemplazar el archivo de entrada
            filename = filename + "x"  # Actualizar el nombre del archivo

        # Determinar el tipo de archivo
        if filename.endswith(".xls"):
            # Convertir .doc a .docx
            docx_file = convert_xls_to_xlsx(input_file)
            input_file = docx_file  # Reemplazar el archivo de entrada
            filename = filename + "x"  # Actualizar el nombre del archivo
        
        # Procesar archivos soportados
        if filename.endswith((".docx", ".pdf", ".pptx", ".xlsx", ".jpg", ".mp3", ".html")):
            # Generar el nombre del archivo .md de salida
            md_filename = os.path.splitext(filename)[0] + ".md"
            md_file = os.path.join(output_dir, md_filename)
            
            # Convertir el archivo a Markdown
            result = markitdown.convert(input_file)
            
            # Guardar el resultado en un archivo .md
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(result.text_content)
            
            print(f"Archivo convertido: {md_file}")
            
            # Eliminar el archivo .docx temporal si se convirtió de .doc
            if filename.endswith(".docx") and input_file.endswith("x"):
                os.remove(input_file)

def process_directory_recursive(input_dir, output_dir):
    """
    Procesa todos los archivos soportados en un directorio y sus subdirectorios, y los convierte a .md.
    """
    # Crear el directorio de salida si no existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Crear una instancia de MarkItDown
    markitdown = MarkItDown()

    # Recorrer todos los archivos en el directorio de entrada y sus subdirectorios
    for root, dirs, files in os.walk(input_dir):
        for filename in files:
            input_file = os.path.join(root, filename)
            relative_path = os.path.relpath(root, input_dir)
            output_subdir = os.path.join(output_dir, relative_path)
            
            # Crear el subdirectorio de salida si no existe
            if not os.path.exists(output_subdir):
                os.makedirs(output_subdir)
            
            # Determinar el tipo de archivo
            if filename.endswith(".doc"):
                # Convertir .doc a .docx
                docx_file = convert_doc_to_docx(input_file)
                input_file = docx_file  # Reemplazar el archivo de entrada
                filename = filename + "x"  # Actualizar el nombre del archivo

            # Determinar el tipo de archivo
            if filename.endswith(".xls"):
                # Convertir .xls a .xlsx
                xlsx_file = convert_xls_to_xlsx(input_file)
                input_file = xlsx_file  # Reemplazar el archivo de entrada
                filename = filename + "x"  # Actualizar el nombre del archivo
            
            # Procesar archivos soportados
            if filename.endswith((".docx", ".pdf", ".pptx", ".xlsx", ".jpg", ".mp3", ".html")):
                # Generar el nombre del archivo .md de salida
                md_filename = os.path.splitext(filename)[0] + ".md"
                md_file = os.path.join(output_subdir, md_filename)
                
                try:
                    # Convertir el archivo a Markdown
                    result = markitdown.convert(input_file)
                    
                    # Guardar el resultado en un archivo .md
                    with open(md_file, "w", encoding="utf-8") as f:
                        f.write(result.text_content)
                    
                    print(f"Archivo convertido: {md_file}")
                except Exception as e:
                    print(f"Error al convertir el archivo {input_file}: {e}")
                
                # Eliminar el archivo .docx temporal si se convirtió de .doc
                if filename.endswith(".docx") and input_file.endswith("x"):
                    os.remove(input_file)

# Ejemplo de uso
if __name__ == "__main__":
    input_directory = "H:\\Unidades compartidas\\PROYECTOS\DEALER3\\10_DISEÑO\\04_CASOS DE USO"  # Cambia esto por la ruta de tu directorio de entrada
    output_directory = "D:\\REPOSITORIOS LOCALES\\DEALER_SERVICES\\07_PY_FILE_TO_MD\\app_data"  # Cambia esto por la ruta de tu directorio de salida
    process_directory_recursive(input_directory, output_directory)
    print(f"Todos los archivos han sido convertidos y guardados en {output_directory}")
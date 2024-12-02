import base64
import icon
from datetime import datetime
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton,QCheckBox, QLineEdit, QFileDialog, QLabel, QInputDialog, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QMenu, QToolButton, QTextEdit
from PyQt5.QtCore import pyqtSignal, QObject, QThread, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QIcon, QPixmap, QMovie
import logging
import unicodedata



def remover_tildes(texto):
    # Normalize the text (NFKD decomposes accented characters into their base form + diacritical marks)
    texto_normalizado = unicodedata.normalize('NFKD', texto)
    # Filter characters ignoring diacritical marks (category Mn)
    texto_sin_acentos = ''.join(c for c in texto_normalizado if not unicodedata.combining(c))
    # Remove any character that is not alphanumeric or an underscore
    texto_sin_acentos = ''.join(c for c in texto_sin_acentos if c.isalnum() or c == '_')
    return texto_sin_acentos

# Function to compare rows in each group with the first row of that group
def compare_group(group):
        # Exclude the grouping column
        comparisons = group.iloc[:, 1:] != group.iloc[0, 1:]  # Ignore the first column ("group")
        return comparisons

def comparar_archivos_excel(archivo1: str, archivo2: str, columna_clave: str, archivo_salida: str, check_cache_archivo1: QCheckBox, check_cache_archivo2: QCheckBox, check_numeros_enteros: QCheckBox, check_eliminar_vacias: QCheckBox, check_ignorar_hora: QCheckBox):
    # Configuración del registro
    logging.basicConfig(filename=archivo_salida+'_log', filemode='a', level=logging.INFO, format='%(message)s\n')
    """Compara dos archivos de Excel y exporta las diferencias a un nuevo archivo."""
    #INIT LOGGING
    logging.info(f"Iniciando una nueva comparación de archivos, en fecha y hora {datetime.now()}")
    estructuraigual = False
    try:
        if check_cache_archivo1.isChecked():
            cache_file1 = archivo1.replace('.xlsx', '_cache.pkl')
            try:
                df1 = pd.read_pickle(cache_file1)
                logging.info(f"{archivo1} cargado desde caché: {cache_file1}")
            except FileNotFoundError:
                df1 = pd.read_excel(archivo1, header=0)
                df1.to_pickle(cache_file1)
                logging.info(f"{archivo1} cargado y guardado en caché: {cache_file1}")
        else:
            df1 = pd.read_excel(archivo1, header=0)

        if check_cache_archivo2.isChecked():
            cache_file2 = archivo2.replace('.xlsx', '_cache.pkl')
            try:
                df2 = pd.read_pickle(cache_file2)
                logging.info(f"{archivo2} cargado desde caché: {cache_file2}")
            except FileNotFoundError:
                df2 = pd.read_excel(archivo2, header=0)
                df2.to_pickle(cache_file2)
                logging.info(f"{archivo2} cargado y guardado en caché: {cache_file2}")
        else:
            df2 = pd.read_excel(archivo2, header=0)
    except Exception as e:
         logging.info(f"Error al leer los archivos de Excel: {e}")
    # Extract the base filenames without path and extension, then get the first 4 characters
    archivo1name = archivo1.split('/')[-1].split('\\')[-1].split('.')[0][:4]
    archivo2name = archivo2.split('/')[-1].split('\\')[-1].split('.')[0][:4]

    
    #CHECK IF THE STRUCTURE OF THE FILES IS THE SAME
    if set(df1.columns) != set(df2.columns):
        logging.info("Los archivos tienen distinta estructura de columnas")
        #GET COLUMN NUMBER OF BOTH FILES
        columnas1 = len(df1.columns)
        columnas2 = len(df2.columns)
        logging.info(f"El {archivo1name} tiene {columnas1} columnas y el {archivo2name} tiene {columnas2} columnas")
    else:
         estructuraigual = True
         logging.info("Los archivos tienen la misma estructura de columnas")
         #GET COLUMN NUMBER OF BOTH FILES
         columnas1 = len(df1.columns)
         logging.info(f"Ambos archivos tienen {columnas1} columnas")
         

    #SET HEADER NAMES CHANGING SPACES AND BREAKLINE TO "_", CHANGE TO UPPER CASE, AND REMOVE SPECIAL CHARACTERS LIKE '´' TO AVOID ERRORS
    df1.columns = df1.columns.str.replace(' ', '_').str.replace('\n', '_').str.upper().map(remover_tildes)
    df2.columns = df2.columns.str.replace(' ', '_').str.replace('\n', '_').str.upper().map(remover_tildes)
    columna_clave = columna_clave.replace(' ','_').replace('\n', '_').upper()
    columna_clave = remover_tildes(columna_clave)

    #EXTRACT TO EXCEL INFORM COMPARING COLUMN NAMES AND INDEX POSITION ORDER BY INDEX OF FILE1
    logging.info("Exportando a Excel la información de las columnas y su posición en el archivo")
    try:
        columnas1 = pd.DataFrame(df1.columns, columns=[archivo1name])
        columnas2 = pd.DataFrame(df2.columns, columns=[archivo2name])
        columnas = pd.concat([columnas1, columnas2], axis=1)
        columnas.to_excel(archivo_salida.replace('.xlsx', '_columnas.xlsx'))
    except Exception as e:
        logging.info(f"Error al exportar la información de las columnas a Excel: {e}")

    #REMOVE UNNAMED COLUMNS
    if check_eliminar_vacias.isChecked():
        #IF HEADER CONTAINS UNNAMED REMOVE THE COLUMN
        logging.info("Eliminando columnas vacías en el archivo maestro")
        for column in df1.columns:
            try:
                if 'Unnamed' in column:
                    df1.drop(columns=[column], inplace=True)
            except Exception as e:
                logging.info(f"Error al eliminar columna vacía '{column}' en el archivo maestro: {e}")
        #IF HEADER CONTAINS UNNAMED REMOVE THE COLUMN
        logging.info("Eliminando columnas vacías en el archivo a comparar")
        for column in df2.columns:
            try:
                if 'Unnamed' in column:
                    df2.drop(columns=[column], inplace=True)
            except Exception as e:
                logging.info(f"Error al eliminar columna vacía '{column}' en el archivo a comparar: {e}")
            
    #CHECK IF THE STRUCTURE OF THE FILES IS THE SAME IN CASE OF DIFFERENT WRITED NAMES
    if set(df1.columns) != set(df2.columns):
        logging.info("Las columnas de los archivos no coinciden en denominación")
        # Log the differences in column names
        columnas_faltantes_1_set = set(df1.columns) - set(df2.columns)
        columnas_faltantes_2_set = set(df2.columns) - set(df1.columns)
        
        try:
            if columnas_faltantes_1_set:
                logging.info(f"Diferencias en las columnas de los archivos:")
                #GROUP UNAMED COLUMNS AND THE OTHERS TO PRINT AGROUPED
                columnas_faltantes_1_grouped = pd.Series(list(columnas_faltantes_1_set)).groupby(pd.Series(list(columnas_faltantes_1_set)).str.contains('UNNAMED')).apply(list)
                logging.info(f"Las columnas que faltan en el {archivo2name} son:")
                #LOG LINE BY LINE
                for columna in columnas_faltantes_1_grouped:
                    logging.info(f"Columna que falta en el {archivo2name}: {columna}")
        except Exception as e:
            logging.info(f"Error al comparar las columnas faltantes en el {archivo2name}: {e}")

        try:
            if columnas_faltantes_2_set:
                logging.info(f"Las columnas que faltan en el {archivo1name} son:")
                #GROUP UNAMED COLUMNS AND THE OTHERS TO PRINT AGROUPED
                columnas_faltantes_2_grouped = pd.Series(list(columnas_faltantes_2_set)).groupby(pd.Series(list(columnas_faltantes_2_set)).str.contains('UNNAMED')).apply(list)
                #LOG LINE BY LINE
                for columna in columnas_faltantes_2_grouped:
                    logging.info(f"Columna que falta en el {archivo1name}: {columna}")
        except Exception as e:
            logging.info(f"Error al comparar las columnas faltantes en el {archivo1name}: {e}")
            
    else:
         logging.info("Las columnas de los archivos tienen la misma denominacion")

    #REPLACE ALL NaN VALUES TO AVOID ERRORS
    logging.info("Reemplazando los valores NaN por cadena vacía para evitar errores")
    df1 = df1.fillna('')
    df2 = df2.fillna('')
    #REPLACE ALL STRING nan VALUES TO AVOID ERRORS
    logging.info("Reemplazando los valores 'nan' por cadena vacía para evitar errores")
    df1 = df1.replace('nan', '')
    df2 = df2.replace('nan', '')

    #ALL DATETIMES TO ONLY DATE
    if check_ignorar_hora.isChecked():
        logging.info("Convirtiendo las fechas a solo fecha")
        for column in df1.columns:
            try:
                if pd.api.types.is_datetime64_any_dtype(df1[column]):
                    df1[column] = df1[column].dt.date
            except Exception as e:
                logging.info(f"Error al convertir las fechas a solo fecha en la columna '{column}' del archivo {archivo1name}: {e}")
            
        for column in df2.columns:
            try:
                if pd.api.types.is_datetime64_any_dtype(df2[column]):
                    df2[column] = df2[column].dt.date
            except Exception as e:
                logging.info(f"Error al convertir las fechas a solo fecha en la columna '{column}' del archivo {archivo2name}: {e}")

    #IF IS NUMERIC CONVERT TO INTEGER
    if check_numeros_enteros.isChecked():
        for column in df1.columns:
            try:
                if pd.api.types.is_numeric_dtype(df1[column]):
                    df1[column] = df1[column].astype(int)
            except Exception as e:
                logging.info(f"Error al convertir los números a enteros en la columna '{column}' del archivo {archivo1name}: {e}")

        for column in df2.columns:
            try:
                if pd.api.types.is_numeric_dtype(df2[column]):
                    df2[column] = df2[column].astype(int)
            except Exception as e:
                logging.info(f"Error al convertir los números a enteros en la columna '{column}' del archivo {archivo2name}: {e}")

    #CONVERT ALL DATA TO STRING TO AVOID ERRORS
    logging.info("Convirtiendo todos los datos a string para evitar errores")
    df1 = df1.astype(str)
    df2 = df2.astype(str)
 
    #REPLACE NaT VALUES TO AVOID ERRORS
    logging.info("Reemplazando los valores 'NaT' por cadena vacía para evitar errores")
    df1 = df1.replace('NaT', '')
    df2 = df2.replace('NaT', '')        

    # Ensure the key column exists in both dataframes
    if columna_clave not in df1.columns or columna_clave not in df2.columns:
        logging.info(f"La columna clave '{columna_clave}' no existe en uno de los archivos.")
        return

    #SET SAME COLUMN INDEX IN BOTH DATAFRAMES
    logging.info("Estableciendo la columna clave como índice en ambos archivos")
    try:
        df1.set_index(columna_clave, inplace=True)
        df2.set_index(columna_clave, inplace=True)
    except Exception as e: 
        logging.info(f"Error al establecer la columna clave seleccionada como índice en ambos archivos: {e}")

    #COMPARE THE KEYS OF BOTH FILES (label_columna_clave)
    logging.info("Comparando las claves de ambos archivos")
    try:
        keys1 = set(df1.index)
        keys2 = set(df2.index)
        keys1_not_in_2 = keys1 - keys2
        keys2_not_in_1 = keys2 - keys1
        if keys1_not_in_2:
            logging.info(f"Claves en {archivo1name} que no están en {archivo2name}: {keys1_not_in_2}")
        if keys2_not_in_1:
            logging.info(f"Claves en {archivo2name} que no están en {archivo1name}: {keys2_not_in_1}")
    except Exception as e:
        logging.info(f"Error al comparar las claves de ambos archivos: {e}")

    #GET ONLY THE ROWS THAT ARE IN BOTH FILES
    logging.info("Obteniendo solo las filas que están en ambos archivos")
    try:
        common_index = df1.index.intersection(df2.index)
        df1 = df1.loc[common_index]
        df2 = df2.loc[common_index]
    except Exception as e:
        logging.info(f"Error al obtener solo las filas que están en ambos archivos: {e}")

    #PREPARE TO COMPARE THE DATA REMOVE FROM DATAFRAME THE COLUMNS THAT ARE NOT IN BOTH FILES
    try:
        logging.info("Eliminando las columnas que no están en ambos archivos")
        common_columns = list(set(df1.columns) & set(df2.columns))
        df1 = df1[common_columns]
        df2 = df2[common_columns]
    except Exception as e:
        logging.info(f"Error al eliminar las columnas que no están en ambos archivos: {e}")

    #COMPARE THE DATA
    logging.info("Comparando los datos de ambos archivos")
    try:
        diff = df1.compare(df2, align_axis=1, keep_shape=False, keep_equal=False, result_names=(archivo1name, archivo2name))
    except Exception as e:
        logging.info(f"Error al comparar los datos de ambos archivos: {e}")

    #CHECK ARCHIVO SALIDA IS A PATH
    if not archivo_salida.endswith('.xlsx'):
        archivo_salida += '.xlsx'

    if not diff.empty:
        try:
            #EXPORT THE DATAFRAME TO AN EXCEL FILE
            diff.to_excel(archivo_salida)
        except Exception as e:
            logging.info(f"Error al exportar las diferencias a Excel: {e}")
        try:
            # Remove columns from diff where both 'self' and 'other' columns are empty
            diff = diff.dropna(axis=1, how='all')
            #EXPORT NAME OF THE OTHER COLUMNS TO A LIST
            columnas = diff.columns
            logging.info("Enumeracion de columnas con diferencias:")
            last_column = ''
            diferencia = 0
            for columna in enumerate(columnas):
                if columna[1][0] != last_column:
                    last_column = columna[1][0]
                    diferencia += 1
                logging.info(f"{diferencia}. Columna con diferencias: {columna}")
                
        except Exception as e:
            logging.info(f"Error al elegir las columnas con diferencias: {e}")
        
        try:
            #EXPORT RESUME OF DIFF TO EXCEL FILE WITH THE TOTAL OF DIFFERENCES PER COLUMN
            resume = diff.apply(lambda x: x.notna().sum())
            resume.to_excel(archivo_salida.replace('.xlsx', '_resumen.xlsx'))
        except Exception as e:
            logging.info(f"Error al exportar el resumen a Excel: {e}")
    else:
        logging.info("No hay diferencias para exportar.")
    

    

class ComparadorExcel(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        """Inicializa los componentes de la interfaz de usuario."""
        layout = QVBoxLayout()
        #layout border
        layout.setContentsMargins(10, 10, 10, 10)
        
        self.label_archivo1 = QLabel('Archivo Maestro:')
        layout.addWidget(self.label_archivo1)
        self.entry_archivo1 = QLineEdit(self)
        layout.addWidget(self.entry_archivo1)
        self.boton_examinar1 = QPushButton('Examinar', self)
        self.boton_examinar1.clicked.connect(lambda: self.seleccionar_archivo(self.entry_archivo1))
        layout.addWidget(self.boton_examinar1)
        
        self.check_cache_archivo1 = QCheckBox('Usar archivo en caché si está disponible', self)
        layout.addWidget(self.check_cache_archivo1)
        
        self.label_archivo2 = QLabel('Archivo a Comparar:')
        layout.addWidget(self.label_archivo2)
        self.entry_archivo2 = QLineEdit(self)
        layout.addWidget(self.entry_archivo2)
        self.boton_examinar2 = QPushButton('Examinar', self)
        self.boton_examinar2.clicked.connect(lambda: self.seleccionar_archivo(self.entry_archivo2))
        layout.addWidget(self.boton_examinar2)

        self.check_cache_archivo2 = QCheckBox('Usar archivo en caché si está disponible', self)
        layout.addWidget(self.check_cache_archivo2)
        
        self.label_columna_clave = QLabel('Índice de Columna Clave:')
        layout.addWidget(self.label_columna_clave)
        self.entry_columna_clave = QLineEdit(self)
        layout.addWidget(self.entry_columna_clave)

        self.check_ignorar_hora = QCheckBox('Ignorar Hora en Fecha y Hora', self)
        self.check_ignorar_hora.setChecked(True)
        layout.addWidget(self.check_ignorar_hora)

        self.check_numeros_enteros = QCheckBox('Considerar números como enteros', self)
        self.check_numeros_enteros.setChecked(True)
        layout.addWidget(self.check_numeros_enteros)

        self.check_eliminar_vacias = QCheckBox('Ignorar columnas vacías', self)
        self.check_eliminar_vacias.setChecked(True)
        layout.addWidget(self.check_eliminar_vacias)

        self.label_archivo_salida = QLabel('Archivo de Salida:')
        layout.addWidget(self.label_archivo_salida)
        self.entry_archivo_salida = QLineEdit(self)
        layout.addWidget(self.entry_archivo_salida)
        self.boton_examinar_salida = QPushButton('Seleccionar', self)
        self.boton_examinar_salida.clicked.connect(self.seleccionar_archivo_salida)
        layout.addWidget(self.boton_examinar_salida)
        
        self.boton_comparar = QPushButton('Comparar', self)
        self.boton_comparar.clicked.connect(self.ejecutar_comparacion)
        layout.addWidget(self.boton_comparar)


    #HABILITY TO TEST WITH PRESELECTED VALUES
        # self.boton_probar = QPushButton('Probar con valores preseleccionados', self)
        # self.boton_probar.clicked.connect(self.probar_valores_preseleccionados)
        # layout.addWidget(self.boton_probar)
        
    #ADD ICONS
        icon_data = base64.b64decode(icon.icon_base64)
        pixmap = QPixmap()
        pixmap.loadFromData(icon_data)
        self.setWindowIcon(QIcon(pixmap))

        self.setLayout(layout)
        self.setWindowTitle('Comparador de Excel')
        self.show()
        
    def probar_valores_preseleccionados(self):
        self.entry_archivo1.setText('./maestro.xlsx')
        self.entry_archivo2.setText('./esclavo.xlsx')
        self.entry_columna_clave.setText('CODIGO')
        self.entry_archivo_salida.setText('./diferencias.xlsx')
        self.check_cache_archivo1.setChecked(True)
        self.check_cache_archivo2.setChecked(True)
        self.check_eliminar_vacias.setChecked(True)
        
    def seleccionar_archivo(self, entry):
        """Abre un cuadro de diálogo para seleccionar un archivo de Excel."""
        ruta_archivo, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo", "", "Archivos de Excel (*.xlsx)")
        if ruta_archivo:
            entry.setText(ruta_archivo)
    
    def seleccionar_archivo_salida(self):
        """Abre un cuadro de diálogo para seleccionar el archivo de salida."""
        ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar Archivo", "", "Archivos de Excel (*.xlsx)")
        if ruta_archivo:
            self.entry_archivo_salida.setText(ruta_archivo)
    
    def ejecutar_comparacion(self):
        """Ejecuta la comparación y maneja los errores."""
        archivo1 = self.entry_archivo1.text()
        archivo2 = self.entry_archivo2.text()
        columna_clave = self.entry_columna_clave.text()
        archivo_salida = self.entry_archivo_salida.text()
        
        if not archivo1 or not archivo2 or not columna_clave or not archivo_salida:
            QMessageBox.critical(self, "Error", "Todos los campos son obligatorios")
            return
        
        try:
            comparar_archivos_excel(archivo1, archivo2, columna_clave, archivo_salida, self.check_cache_archivo1, self.check_cache_archivo2, self.check_numeros_enteros, self.check_eliminar_vacias, self.check_ignorar_hora)
            QMessageBox.information(self, "Éxito", f"Diferencias exportadas a {archivo_salida}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == '__main__':
    app = QApplication([])
    ex = ComparadorExcel()
    app.exec_()
'''
Sentiment Analyzer for Course Reviews in the AESCULAP ACADEMY platform.

Author : Daniel Malváez
Version : 1.0.0
Date : 2023-10-14

Catching up the notebook:
    1. Install the required libraries:
        $ pip3 install -r requirements.txt
    2. Run the script:
        $ python3 AnalisisSentimientos.py Excel_File Course_Name
    3. The script will generate a PNG image with the results and a PDF report.
    4. The PNG image will be saved as 'Opiniones.png'
    5. The PDF report will be saved as 'Reporte Higiene Manos.pdf'
    6. The script will print the total score of the course in stars.
    7. The script will print the PDF path.
'''
# ------------------------------------
# Important libraries for the analysis
# ------------------------------------
# Data Manipulation & Visualization
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Text Preprocessing
from deep_translator import GoogleTranslator

# Sentiment Analysis
import nltk
from nltk.sentiment import SentimentIntensityAnalyzer
from nltk.corpus import stopwords
nltk.download('stopwords')
from unidecode import unidecode

# Report Generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle, TA_JUSTIFY

# Regular Expressions
import re

# System
import sys

# DateTime
import datetime

# ------------------------------
# Defining the functions to use
# ------------------------------
def quitar_caracteres_especiales(texto):
    # Regex for special characters
    texto_limpio = re.sub(r'[^a-zA-Z\sáéíóúñ]', '', texto)
    return texto_limpio

def traducir(opinion):
    # Translate the opinion into english
    opinion_trad = GoogleTranslator(source='es', target='en').translate(opinion)
    opinion_trad = unidecode(opinion_trad)
    opinion_trad = opinion_trad.lower()
    opinion_trad = re.sub(r'[^\w\s]', '', opinion_trad)
    return opinion_trad

def quitar_stops(texto, stop_words = stopwords.words('english')):
    # Remove stopwords
    texto_limpio = ' '.join([word for word in texto.split() if word not in stop_words])
    return texto_limpio

def main():
    # Verifying the arguments
    if len(sys.argv) < 3 or len(sys.argv) > 3:
        print('Error - Utiliza el siguiente orden de argumentos')
        print('Uso: $ python3 AnalisisSentimientos.py Excel_File Course_Name')
        sys.exit(1)
    
    print('Cargando recursos necesarios...')
    
    sys.stdout.reconfigure(line_buffering=True)
    nltk.download('stopwords')
    nltk.download('vader_lexicon')

    print('Recursos cargados.')

    # Obtaining the arguments
    excel_file = sys.argv[1]
    course_name = sys.argv[2]

    # Reading the excel file
    opinions = pd.read_excel(excel_file, sheet_name = 'Opiniones')

    # Obtaining the opinions
    pos_ops = [2+(3*x) for x in range(788)]
    opinions_only = list([opinions.iloc[x]['OPINIONES'] for x in pos_ops])

    # --------------------------
    # Preprocessing the opinions
    # --------------------------
    opinions_only = [quitar_caracteres_especiales(str(x).strip().lower()) for x in opinions_only]
    df = pd.DataFrame(opinions_only, columns = ['Opiniones'])
    df[df['Opiniones'] == 'nan'] = np.nan
    df[df['Opiniones'] == ''] = np.nan
    df.dropna(inplace = True)
    df.reset_index(drop = True, inplace = True)
    opinions_only = df['Opiniones'].tolist()
    
    print('traduciendo opiniones...')
    df_trad = df.applymap(traducir)

    print('limpiando opiniones...')
    df_cleaned = df_trad.applymap(quitar_stops)

    # --------------------------
    # Sentiment Analysis
    # --------------------------
    sia = SentimentIntensityAnalyzer()
    
    # Obtaining the opinions as a list
    opinions_cleaned = df_cleaned['Opiniones'].tolist()

    # Polarity scores
    polarity_scores = [sia.polarity_scores(review)['compound'] for review in opinions_cleaned]

    # Defining the ranges for the polarity scores
    rangos = np.linspace(-1,1, 4)
    etiquetas = ['Mala', 'Regular', 'Buena']

    # Obtaining the ranges for the polarity scores
    polarity_ranges = pd.cut(polarity_scores, bins=rangos, labels=etiquetas)

    # --------------------------
    # Data Visualization
    # --------------------------
    colores = ['#FC540B', '#FCCD0B', '#35D773']

    # Style
    sns.set(style='whitegrid')
    sns.set_palette(colores)

    # Figure
    ax = sns.countplot(x=polarity_ranges)
    for p in ax.patches:
        ax.annotate(format(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha = 'center', va = 'center', xytext = (0, 5), textcoords = 'offset points')

    # Labels and Title
    plt.xlabel('Clasificación', fontweight='bold', labelpad=4, fontfamily='Arial')
    plt.ylabel('Frecuencia', fontweight='bold', labelpad=4, fontfamily='Arial')
    plt.title(f'Comportamiento de las opiniones del curso {course_name}',
              fontsize=16, fontweight='bold', pad=20, loc='center', fontfamily='Arial')

    # Save the figure and show
    plt.savefig('Opiniones.png', dpi=300, bbox_inches='tight', format='png')

    # --------------------------
    # Generating the PDF report
    # --------------------------
    
    # Number of opinions by type
    total_opinions = len(df_cleaned)
    good_ops = polarity_ranges.value_counts()['Buena']
    regular_ops = polarity_ranges.value_counts()['Regular']
    bad_ops = polarity_ranges.value_counts()['Mala']

    # Calculate the overall score (5 starss)
    score_bad = 1
    score_regular = 3
    score_good = 5
    overall_score = (bad_ops * score_bad) + (regular_ops * score_regular) + (good_ops * score_good)

    # Create a PDF file
    pdf_path = f'Reporte del curso {course_name}.pdf'
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)

    # Create a list to hold the story (elements to be added to the PDF)
    story = []
    
    # Define styles
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    title_style = styles['Title']
    heading_style = styles['Heading2']
    # Crear un estilo personalizado con alineación justificada
    custom_style = ParagraphStyle(name="CustomStyle", parent=styles["Normal"], alignment=TA_JUSTIFY)

    # Add the current date and author
    current_date = datetime.datetime.now().strftime('%Y-%m-%d')
    author = 'Axel Daniel Malváez Flores' 
    date_paragraph = Paragraph(f'<b>Fecha del Informe:</b> {current_date}', normal_style)
    author_paragraph = Paragraph(f'<b>Autor:</b> {author}', normal_style)

    story.append(date_paragraph)
    story.append(author_paragraph)
    story.append(Spacer(1, 12))

    # Add a title to the PDF
    title = Paragraph(f'Informe de Análisis de Sentimientos para el Curso {course_name}', title_style)
    story.append(title)
    story.append(Spacer(1, 12))
    
    # Introduction Seccion
    introduction = f'''
    Este informe de análisis de sentimientos examina las opiniones enviadas por estudiantes (participantes) del curso {course_name}
    dentro de la plataforma de la Fundación Academia Aesculap México. Esta plataforma se utiliza para impartir cursos de capacitación
    en línea para profesionales de la salud. El objetivo de este informe es proporcionar una visión general de las opiniones de los
    estudiantes y brindar recomendaciones para mejorar la calidad del curso. Link plataforma :<u>http://academiaaesculap.eadbox.com/</u>
    '''
    introduction_paragraph = Paragraph(introduction, custom_style)
    story.append(introduction_paragraph)
    story.append(Spacer(1, 12))
    
    # Methodology Seccion
    methodology_title = 'Metodología'
    methodology = f'''
    Para llevar a cabo este análisis, se recopilaron opiniones de los estudiantes. Estas opiniones se tradujeron al inglés y se
    limpiaron para eliminar caracteres especiales y palabras vacías. Luego, se analizaron las opiniones utilizando el paquete
    NLTK de Python. Este paquete utiliza un diccionario de palabras para asignar un puntaje de polaridad a cada palabra. El
    puntaje de polaridad de cada palabra se promedia para obtener un puntaje de polaridad para cada opinión. Finalmente, se
    agruparon las opiniones en tres categorías: 'Buenas', 'Regulares' y 'Malas'. Las opiniones 'Buenas' tienen un puntaje de
    polaridad mayor o igual a 0.3, las opiniones 'Regulares' tienen un puntaje de polaridad mayor o igual a -0.3 y menor a 0.3,
    y las opiniones 'Malas' tienen un puntaje de polaridad menor a -0.3.
    '''
    methodology_paragraph = Paragraph(methodology_title, heading_style)
    methodology_content = Paragraph(methodology, custom_style)
    story.append(methodology_paragraph)
    story.append(methodology_content)
    story.append(Spacer(1, 12))
    
    # Results Seccion
    results_title = 'Resultados'

    original_width = 640
    original_height = 480
    scale_factor = 0.5
    
    new_width = int(original_width * scale_factor)
    new_height = int(original_height * scale_factor)
    image = Image('Opiniones.png', width=new_width, height=new_height)
    
    results = f'''
    En el siguiente análisis, se presenta una representación gráfica de la distribución de
    opiniones en el curso. En total, se han evaluado {total_opinions} opiniones. De estas, {good_ops} opiniones
    son calificadas como 'Buenas', {regular_ops} como 'Regulares' y {bad_ops} como 'Malas'.
    '''    
    results_paragraph = Paragraph(results_title, heading_style)
    results_content = Paragraph(results, custom_style)
    story.append(results_paragraph)
    story.append(image)
    story.append(results_content)
    story.append(Spacer(1, 12))
    
    results2 = f'''
    Si evaluáramos el curso en una escala de 0 a 5 estrellas, el puntaje promedio sería de
    {overall_score / total_opinions:.2f} estrellas.
    '''
    results2_paragraph = Paragraph(results2, custom_style)
    story.append(results2_paragraph)
    story.append(Spacer(1, 12))
    
    # Obtenemos un ejemplo de cada tipo de comentario
    buena_index = -1
    regular_index = -1
    mala_index = -1

    example_good_comment = ''
    example_regular_comment = ''
    example_bad_comment = ''
    
    elemento_buscado = 'Buena'
    if elemento_buscado in list(polarity_ranges):
        buena_index = list(polarity_ranges).index(elemento_buscado)

    elemento_buscado = 'Regular'
    if elemento_buscado in list(polarity_ranges):
        regular_index = list(polarity_ranges).index(elemento_buscado)

    elemento_buscado = 'Mala'
    if elemento_buscado in list(polarity_ranges):
        mala_index = list(polarity_ranges).index(elemento_buscado)
        
    idx = [buena_index, regular_index, mala_index]
    for i in range(len(idx)):
        if idx[i] != -1 and i == 0:
            example_good_comment = opinions_only[idx[i]]
        elif idx[i] != -1 and i == 1:
            example_regular_comment = opinions_only[idx[i]]
        elif idx[i] != -1 and i == 2:
            example_bad_comment = opinions_only[idx[i]]
        else:
            continue

    story.append(Paragraph('Ejemplos de Comentarios:', custom_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f'<b>Comentario Bueno:</b> {example_good_comment}', custom_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f'<b>Comentario Regular:</b> {example_regular_comment}', custom_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f'<b>Comentario Malo:</b> {example_bad_comment}', custom_style))
    
    # Recomendaciones
    recommendations_title = 'Recomendaciones'
    recommendations = '''
    Considerando los resultados del análisis de sentimientos, se recomienda implementar un enfoque
    equilibrado para abordar las opiniones críticas ('Malas') y mantener la satisfacción de los
    estudiantes que han expresado opiniones positivas ('Buenas'). Esto implica mejorar las áreas que
    generan opiniones negativas y mantener o reforzar las prácticas que resultan en opiniones positivas.
    '''
    recommendations_paragraph = Paragraph(recommendations_title, heading_style)
    recommendations_content = Paragraph(recommendations, custom_style)
    story.append(recommendations_paragraph)
    story.append(recommendations_content)
    
    # Build the PDF
    doc.build(story)
    print(f'PDF generado como : {pdf_path}')
    
if __name__ == '__main__':
    main()
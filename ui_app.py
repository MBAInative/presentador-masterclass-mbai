import streamlit as st
import os
import sys

# Ajuste temporal del path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from ai_presenter import AIPresenterPipeline

st.set_page_config(page_title="MBAI Native AI Presenter", layout="wide")

st.title("🚀 MBAI Native - Generador de Presentaciones Corporativo")
st.markdown("Transforma cualquier tópico en una Masterclass visual y redactada al instante, en el idioma que desees.")

with st.sidebar:
    st.header("🔑 Configuración de APIs")
    openrouter_key = st.text_input("OpenRouter API Key (Motor Lógico)", type="password")
    
    st.header("🌍 Idioma y Opciones Multimedia")
    language_choice = st.selectbox(
        "Idioma del PPTX (Textos y Guiones)",
        options=["Español", "Inglés", "Francés", "Alemán", "Portugués", "Italiano"]
    )
    
    voice_map = {
        "Español": {"es-ES-AlvaroNeural": "Hombre (España)", "es-ES-ElviraNeural": "Mujer (España)", "es-MX-JorgeNeural": "Hombre (México)", "es-MX-DaliaNeural": "Mujer (México)"},
        "Inglés": {"en-US-ChristopherNeural": "Hombre (USA)", "en-US-AriaNeural": "Mujer (USA)", "en-GB-RyanNeural": "Hombre (UK)", "en-GB-SoniaNeural": "Mujer (UK)"},
        "Francés": {"fr-FR-HenriNeural": "Hombre (Francia)", "fr-FR-DeniseNeural": "Mujer (Francia)"},
        "Alemán": {"de-DE-KillianNeural": "Hombre (Alemania)", "de-DE-AmalaNeural": "Mujer (Alemania)"},
        "Portugués": {"pt-PT-DuarteNeural": "Hombre (Portugal)", "pt-PT-RaquelNeural": "Mujer (Portugal)", "pt-BR-AntonioNeural": "Hombre (Brasil)", "pt-BR-FranciscaNeural": "Mujer (Brasil)"},
        "Italiano": {"it-IT-DiegoNeural": "Hombre (Italia)", "it-IT-ElsaNeural": "Mujer (Italia)"}
    }
    
    available_voices = voice_map[language_choice]
    inverted_voices = {v: k for k, v in available_voices.items()}
    voice_display_choice = st.selectbox("Locutor Neuronal", options=list(inverted_voices.keys()))
    selected_voice_id = inverted_voices[voice_display_choice]
    
    tts_speed = st.slider("Velocidad de Narración TTS (%)", min_value=-50, max_value=50, value=0, step=5, help="0% es la velocidad neuronal humana idónea nativa.")
    
    generate_tts = st.checkbox("💾 Exportar Mp3 Narración (Generar audios separados con locutores TTS de cada diapositiva)", value=True)
    
    st.header("🤖 Selección de Capacidades")
    
    pdf_upload = st.file_uploader("📑 MODO ESTRICTO: Subir documento PDF base (Opcional).", type=["pdf"], help="Si adjuntas un PDF, la IA no investigará en internet, todo el contenido de la clase y las gráficas se obtendrán milimétricamente destripando tu documento PDF.")
    
    model_choice = st.selectbox(
        "Cerebro Lógico Principal (vía API OpenRouter)",
        options=[
            "openai/gpt-4o-mini",
            "google/gemini-2.5-flash",
            "anthropic/claude-3-haiku",
            "meta-llama/llama-3-8b-instruct"
        ],
        index=0
    )
    
    image_source = st.radio(
        "Origen de Ilustraciones (Paso 3)",
        options=[
            "Búsqueda Web Autónoma de Wikipedia (Gratis)", 
            "IA DALL-E 3 Limitada (Requiere API OpenAI)", 
            "Extraer Gráficos Interiores del PDF (Solo si subes PDF)"
        ],
        index=0
    )
    
    st.header("🏢 Branding Corporativo")
    footer_text = st.text_input("Texto de Pie de Página (Esquina Inferior Derecha)", "MBAI NATIVE")
    
    openai_key = ""
    if "DALL-E" in image_source:
        openai_key = st.text_input("OpenAI API Key (Exclusiva para DALL-E 3 Imágenes)", type="password")

topic = st.text_input("Tema de la Presentación", "El impacto geopolítico de la IA Convencional")
num_slides = st.slider("Cantidad Diapositivas Demandadas", min_value=1, max_value=30, value=5)

if st.button("✨ Generar Masterclass Magistral", use_container_width=True):
    with st.spinner(f"Iniciando flujo de orquestación autónomo para generar {num_slides} páginas sobre '{topic[:40]}' en {language_choice}..."):
        if openai_key:
            os.environ["OPENAI_API_KEY"] = openai_key
        if openrouter_key:
            os.environ["OPENROUTER_API_KEY"] = openrouter_key
        
        
        # Guardado temporal del PDF
        target_pdf_path = None
        if pdf_upload:
            os.makedirs("assets", exist_ok=True)
            target_pdf_path = os.path.join("assets", "uploaded_doc.pdf")
            with open(target_pdf_path, "wb") as f:
                f.write(pdf_upload.getbuffer())
        
        os.environ["OR_MODEL_CHOICE"] = model_choice
        
        # Mapeo de la opcion de imagen
        source_param = "web"
        if "DALL-E" in image_source:
            source_param = "dalle"
        elif "PDF" in image_source:
            source_param = "pdf"
            
        try:
            pipeline = AIPresenterPipeline()
            import re, time
            timestamp = int(time.time())
            safe_topic = re.sub(r'[\\/*?:"<>|]', "", topic[:40])
            pptx_name = f"presentacion_{safe_topic.replace(' ', '_')}_{timestamp}.pptx"
            
            pipeline.run(
                topic=topic,
                num_slides=num_slides,
                upload_gws=False,
                image_source=source_param,
                forced_name=pptx_name,
                language=language_choice,
                generate_tts=generate_tts,
                pdf_path=target_pdf_path,
                footer_text=footer_text,
                tts_speed=tts_speed,
                tts_voice=selected_voice_id
            )
            
            full_path = os.path.abspath(pptx_name)
            
            # Logica para agrupar paquete si TTS está activado
            download_path = full_path
            download_name = pptx_name
            mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            download_label = "📥 Descargar Masterclass PowerPoint (.pptx)"
            
            if generate_tts and os.path.exists("assets/audio"):
                import zipfile
                zip_name = f"PAQUETE_{pptx_name.replace('.pptx', '')}.zip"
                zip_path = os.path.abspath(zip_name)
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    # Añadir pptx
                    if os.path.exists(full_path):
                        zipf.write(full_path, arcname=pptx_name)
                    # Añadir audios
                    for root, dirs, files in os.walk("assets/audio"):
                        for file in files:
                            # Solo empaquetamos audios relativos a esta sesion via ordenación o todos generados recientes
                            if file.endswith(".mp3"):
                                zipf.write(os.path.join(root, file), arcname=os.path.join("audio", file))
                
                download_path = zip_path
                download_name = zip_name
                mime_type = "application/zip"
                download_label = "📦 Descargar Paquete ZIP (PPTX Multimedia + Pistas de Audio MP3)"
                
            
            st.success("¡Pipeline exitoso! Su Masterclass con Notas de Orador integradas ha sido renderizada exhaustivamente.")
            
            if os.path.exists(download_path):
                with open(download_path, "rb") as file:
                    st.download_button(
                        label=download_label,
                        data=file,
                        file_name=download_name,
                        mime=mime_type,
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"❌ Ocurrió un error bloqueante en la fase de investigación / diseño gráfico: {e}")

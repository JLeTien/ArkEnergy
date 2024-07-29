import google.generativeai as genai
import os
from AI_module.apikey import gemini_key

def summarise_gemini(input_text):
    genai.configure(api_key=gemini_key)
    model = genai.GenerativeModel('gemini-1.5-flash')
    prompt = f"summarise the following in 5 dotpoints. Place numbers in front of each dotpoint: {input_text}"
    response = model.generate_content(prompt)
    
    # Ensure that the response is plain text
    return response.text.strip()
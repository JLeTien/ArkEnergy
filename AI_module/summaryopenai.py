import openai
from AI_module.apikey import *

openai.api_key = api_key

def summarise(input_text: str,engine_choice ="gpt-3.5-turbo-instruct", max_tok = 3000, temp =0.7):
    """
    engine_choice : set to a default value of gpt-3.5 turbo 
    max_token: the max number of tokes the model will respond with. Does not limit 
    how many tokens you can feed the engin 
    temp: a value from 0-1 that dictates the creativity of the model. 0.7 tend to
    be the best

    """
    response = openai.Completion.create(
    engine=engine_choice,  # Updated model name
    prompt=f"summarise the following in 5 dotpoints.Place numbers infront of each dotpoint.{input_text}",
    max_tokens=max_tok,
    temperature=temp)

    return response.choices[0].text.strip()
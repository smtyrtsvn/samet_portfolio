import os
import openai

openai.api_key = os.getenv("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")

response = openai.Completion.create(
  model="text-davinci-003",
  prompt="",
  temperature=0.7,
  max_tokens=256,
  top_p=1,
  frequency_penalty=0,
  presence_penalty=0
)

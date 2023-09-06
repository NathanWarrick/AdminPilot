import whisper
import openai
import requests

openai.api_key = "API Here"

model = whisper.load_model("tiny.en")
result = model.transcribe(
    r"workspace\clients\openai\whisper\test_files\transcribing_1.mp3"
)
response = result["text"]
print(response)

# summary = openai.Completion.create(
#   model="text-davinci-003",
#   prompt="Summarise the following transription in a form that would be suitable for a councils minutes, complete with enough detail that somoene that didn't attend the meeting would understand everything that happened:" + response,
#   max_tokens=300,
#   temperature=0,
#   n = 1
# )

# print("==========================================================================================================================================================+==========")
# for choice in summary['choices']:
#     print(choice['text'])


API_URL = "https://api-inference.huggingface.co/models/nomic-ai/gpt4all-j"
headers = {"Authorization": "Bearer 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"}


def query(payload):
    response = requests.post(API_URL, headers=headers, json=payload)
    return response.json()


output = query(
    {
        "inputs": "Summarise the following transription in a form that would be suitable for a councils minutes, complete with enough detail that somoene that didn't attend the meeting would understand everything that happened:"
        + response,
    }
)
print(output)

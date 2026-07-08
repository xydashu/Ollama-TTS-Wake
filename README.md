Setup:

Open Powershell and run the command 'ollama'
Select the model you want to run and copy its full name e.g 'nexusriot/deepseek-r1-abliterated:8b' 
You can also find the name from the ollama website

In the python file set the self.model_name to the name of your model.

open powershell again and run 'ollama serve'

If it returns 'Error: listen tcp 127.0.0.1:11434: bind: Only one usage of each socket address (protocol/network address/port) is normally permitted.' then you're good to go.

You will have to go to your CMD and run 'pip install numpy sounddevice webrtcvad openai-whisper ollama piper-tts' as these are the dependencies
Then download VOSK from https://alphacephei.com/vosk/models and select the model named 'vosk-model-small-en-us-0.15' it is 40 mb

Once these are done and youve opened the python file open the settings and set everything correctly. ensure you run the vosk test and set the mic.

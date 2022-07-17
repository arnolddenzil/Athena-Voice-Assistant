﻿# Athena-Voice-Assistant

Installation :-
- Clone the repository.
- Copy the PyAudio file to clipboard.
- Open project in any IDE.
- Check where the PyAudio file should be present inside the requirements.txt and modify that line according to your computer(PyAudio @ file:///C:/Users/Arnold/Downloads/PyAudio-0.2.11-cp310-cp310-win_amd64.whl).
- Go to the downloads as specified in the path and paste the PyAudio file.
- Note that you can change the path according to your needs. Just make sure that you keep the PyAudio file in that specified path.

- Next create a new environment variable for the project.
- For that open terminal inside the IDE and activate the environment manually.(In windows if environment variable is 'env' , activate it by entering 'env/Scripts/activate').
  - If there is an error “Running Scripts Is Disabled on This System”, then fix this by opening PowerShell as administrator and run :    
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  - Type 'y' enter.
- Inside terminal after activating environment variable type and enter : pip install -r requirements.txt
- To run : python main.py (enter)

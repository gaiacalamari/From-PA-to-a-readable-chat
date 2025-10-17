# From-PA-to-a-readable-chat
I developed this code to make PA's printed PDF chat reports easier to read, without loosing any information.

Phase 0: Preliminary step
-	Export the conversation of interest from the Cellebrite Physical Analyzer software in Excel format, naming the file 'chat.xlsx':
    - Open the file and delete the unnecessary fields, but leave the following columns and make sure that are written in the following ways: 'Timestamp-Time', 'Direction', 'Participants', 'Body' and 'attachments' 
    - Inside the 'instant_messages\WhatsApp\1' folder you can find the files attached to the chat, including all the audios. 
    - Pay attention: it is important to change the name of the column ‘Attachments #1’ in ‘attachements’. Rename all of them in the same way.

Step 1: Conversation from opus to wav
-	Inside the folder containing the .bat file (that converts the audio from .opus to .wav), create a folder called 'opus' and insert all the audio files exported from the chat (in 'instant_messages\WhatsApp\1'). 
-	Then create a destination folder called 'My_audio', which will contain all the audio converted to .wav

  <img width="178" height="102" alt="immagine" src="https://github.com/user-attachments/assets/5d613d30-1679-4121-aebc-48d41f96174c" />
  
-	At this point double click on the .bat file and wait
    - After a few seconds all the files are converted in ‘My_audio’ folder.
 	
Step 2: Transcription into a word document
-	In the same folder, insert the Python code "trascription_all.py"
 
-	Create the output folder called 'Trascrizione'
-	Via cmd, navigate to the location of the Python file and run it:
      o	Once complete, all audio files are transcribed into the trascrizione.docx file;
      o	Note: it performs approximately 450 transcriptions per hour.
      o	Note: a docx file is saved for every 100 transcriptions. If something happens (sudden PC shutdown, reboot caused by an update…), by restarting the script, you’ll start from the last transcription saved into that temporary file. At the end of the transcriptions,            you’ll have a single docx.

Step 3: Run the chatcreator.py


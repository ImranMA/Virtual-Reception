<h3>Virtual Reception (Bot)</h3>

Virtual reception Bot is virtual help desk assistant . The idea is simple, if someone goes to any business/office reception the virtual reception will 
greet the user, signs up the user and takes picture. The Bot can be trained to pick up different user intentions and context by training the LUIS Model. 
I have used different samples from microsoft and combined them to make this application. 

<h3>Tech Stack</h3>

I have used Microsoft Cognitive services , FaceAPI, Language Understanding (LUIS) , Bing Speech API (speech to Text) and Speech.dll(Text to speech)
You need to replace the following in App.config 


<h3>Training the model</h3>
This is very simple application wtih limited scope. Since the Bot leverages LUIS to understand the context of user speech , the model can be improved and trained to get better with time.


<h3>How it works</h3>
When user speaks to the Bot , The Bot converts speech to text and analyse the user intent using LUIS API. Once the intent is cleared
and it gets to a point where the user has to be signed up ,the camera gets operational and use FaceAPI to make sure the user is in the frame and 
similing :) You need to do the following to make it work

1) Download the code and simply compile. 
    
2) Deploy the LUIS  (available on rool Virtual Reception.json). You need account to create LUIS app and deploy . 

3) Update the following config with valid Keys and URLs. (You need azure subscription)

```    
     <!-- The character '&' in the LUIS endpoint URL must be replaced by the XML entity reference '&amp;' in order to make the XML parser correctly interpret the file -->
    <!-- For example, https://xxxx&verbose=true&timezoneOffset=0&q= must be changed to https://xxxx&amp;verbose=true&amp;timezoneOffset=0&amp;q=-->
    <add key="LuisEndpointUrl" value="" />     
    
    <add key="AuthenticationUri" value="<End Point Bing Speech>/issueToken" />
    
    <!-- Bing Speech API Key -->
    <add key="subsKeySpeech" value="" />
    
    <!-- Face API Key -->
    <add key="FaceAPIKey" value="" />
    
     <!-- Face API endpoint -->
    <add key="FaceAPIEndPoint" value="<End Point Face API>" />  
    ```
    
    
    
    

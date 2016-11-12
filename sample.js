/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function() {
    "use strict";
    
    /* Recorder Stuff */
    var audio_context;
    var recorder;

    /* UI components */
    var button;
    var buttonEnd;
    var ul;

    var interval;

    /* Set up the Recorder */

    window.onload = function init() {
        try {
          // webkit shim
          window.AudioContext = window.AudioContext || window.webkitAudioContext;
          navigator.getUserMedia = navigator.getUserMedia || navigator.webkitGetUserMedia;
          window.URL = window.URL || window.webkitURL;
          console.log("windows", window);
          audio_context = new AudioContext;

          console.log("audio context", audio_context);
          
          console.log('Audio context set up.');
          console.log('navigator.getUserMedia ' + (navigator.getUserMedia ? 'available.' : 'not present!'));
        } catch (e) {
          console.log("ERROR", e);
        }
        
        navigator.getUserMedia({audio: true}, startUserMedia, function(e) {
          console.log('No live audio input: ' + e);
        });
    };

    // The initialize function is run each time the page is loaded.
    // Office.initialize = function (reason) {
        $(document).ready(function () {
            // Use this to check whether the new API is supported in the Word client.
            // if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
                console.log('This code is using Word 2016 or greater.');

                // Init UI Components
                button = $('#start');
                buttonEnd = $('#end');
                ul = $('#recordingslist');

                // on click listeners
                button.click(function() {
                  button.html('Save');
                  startListening();
                });

                buttonEnd.click(function() {
                  button.html('Start Listening');
                  buttonEnd.html('End Listening');
                  endListening();
                });
            // }
        });
    // };

    function startUserMedia(stream) {
      var input = audio_context.createMediaStreamSource(stream);
      console.log('Media stream created.');

      // Uncomment if you want the audio to feedback directly
      //input.connect(audio_context.destination);
      //__log('Input connected to audio context destination.');
      
      recorder = new Recorder(input);
      console.log('Recorder initialised.');
    }

    function startListening() {
      recorder && recorder.record();

      interval = setInterval(function() {
        createDownloadLink();

        setTimeout(function(){ 
          console.log("Restart the recording"); 
          recorder.clear();
          recorder && recorder.record();
        }, 1000);
      }, 6000);

      button.disabled = true;
      console.log('Recording...');
    }


    function endListening() {
      recorder && recorder.stop();
      buttonEnd.disabled = true;
      console.log('Stopped recording.');
      clearInterval(interval);
      // create WAV download link using audio data blob
      // createDownloadLink();

      recorder.clear();
    }
    
    // create WAV download link using audio data blob
    function createDownloadLink() {

      recorder && recorder.exportWAV(function(blob) {
        var url = URL.createObjectURL(blob);
        var li = document.createElement('li');
        var au = document.createElement('audio');
        var hf = document.createElement('a');
        console.log("url", url);
        
        au.controls = true;
        au.src = url;
        console.log("au", au);
        hf.href = url;
        hf.download = new Date().toISOString() + '.wav';

        hf.innerHTML = hf.download;
        li.appendChild(au);
        li.appendChild(hf);

      });
   }

   // EndPoint https://speech.platform.bing.com/synthesize
   function callTextToSpeechAPI() {
    var request = new XMLHttpRequest();
      request.open("POST", 'https://speech.platform.bing.com/synthesize', true);
      request.onreadystatechange = function() { 
          if (request.readyState === 4 && request.status === 200)
              console.log(request.responseText);
      }
      request.send(null);

   }
})();

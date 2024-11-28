// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.

// Global objects
let speechRecognizer;
let avatarSynthesizer;
let peerConnection;
let messages = [];
let messageInitiated = false;
let dataSources = [];
const sentenceLevelPunctuations = ['.', '?', '!', ':', ';', '。', '？', '！', '：', '；'];
let enableQuickReply = false;
const quickReplies = ['Let me take a look.', 'Let me check.', 'One moment, please.'];
const byodDocRegex = new RegExp(/\[doc(\d+)\]/g);
let isSpeaking = false;
let spokenTextQueue = [];
let sessionActive = false;
let lastSpeakTime;

// Connect to avatar service
async function connectAvatar() {
    const cogSvcRegion = document.getElementById('region').value;
    const cogSvcSubKey = document.getElementById('subscriptionKey').value;
    if (!cogSvcSubKey) {
        alert('Please fill in the subscription key of your speech resource.');
        return;
    }

    const privateEndpointEnabled = document.getElementById('enablePrivateEndpoint').checked;
    const privateEndpoint = document.getElementById('privateEndpoint').value.slice(8);
    if (privateEndpointEnabled && !privateEndpoint) {
        alert('Please fill in the Azure Speech endpoint.');
        return;
    }

    let speechSynthesisConfig;
    if (privateEndpointEnabled) {
        speechSynthesisConfig = SpeechSDK.SpeechConfig.fromEndpoint(new URL(`wss://${privateEndpoint}/tts/cognitiveservices/websocket/v1?enableTalkingAvatar=true`), cogSvcSubKey);
    } else {
        speechSynthesisConfig = SpeechSDK.SpeechConfig.fromSubscription(cogSvcSubKey, cogSvcRegion);
    }
    speechSynthesisConfig.endpointId = document.getElementById('customVoiceEndpointId').value;

    const talkingAvatarCharacter = document.getElementById('talkingAvatarCharacter').value;
    const talkingAvatarStyle = document.getElementById('talkingAvatarStyle').value;
    const avatarConfig = new SpeechSDK.AvatarConfig(talkingAvatarCharacter, talkingAvatarStyle);
    avatarConfig.customized = document.getElementById('customizedAvatar').checked;
    avatarSynthesizer = new SpeechSDK.AvatarSynthesizer(speechSynthesisConfig, avatarConfig);
    avatarSynthesizer.avatarEventReceived = (s, e) => {
        const offsetMessage = e.offset === 0 ? '' : `, offset from session start: ${e.offset / 10000}ms.`;
        console.log(`Event received: ${e.description}${offsetMessage}`);
    };

    const speechRecognitionConfig = SpeechSDK.SpeechConfig.fromEndpoint(new URL(`wss://${cogSvcRegion}.stt.speech.microsoft.com/speech/universal/v2`), cogSvcSubKey);
    speechRecognitionConfig.setProperty(SpeechSDK.PropertyId.SpeechServiceConnection_LanguageIdMode, "Continuous");
    const sttLocales = document.getElementById('sttLocales').value.split(',');
    const autoDetectSourceLanguageConfig = SpeechSDK.AutoDetectSourceLanguageConfig.fromLanguages(sttLocales);
    speechRecognizer = SpeechSDK.SpeechRecognizer.FromConfig(speechRecognitionConfig, autoDetectSourceLanguageConfig, SpeechSDK.AudioConfig.fromDefaultMicrophoneInput());

    const azureOpenAIEndpoint = document.getElementById('azureOpenAIEndpoint').value;
    const azureOpenAIApiKey = document.getElementById('azureOpenAIApiKey').value;
    const azureOpenAIDeploymentName = document.getElementById('azureOpenAIDeploymentName').value;
    if (!azureOpenAIEndpoint || !azureOpenAIApiKey || !azureOpenAIDeploymentName) {
        alert('Please fill in the Azure OpenAI endpoint, API key, and deployment name.');
        return;
    }

    dataSources = [];
    if (document.getElementById('enableOyd').checked) {
        const azureCogSearchEndpoint = document.getElementById('azureCogSearchEndpoint').value;
        const azureCogSearchApiKey = document.getElementById('azureCogSearchApiKey').value;
        const azureCogSearchIndexName = document.getElementById('azureCogSearchIndexName').value;
        if (!azureCogSearchEndpoint || !azureCogSearchApiKey || !azureCogSearchIndexName) {
            alert('Please fill in the Azure Cognitive Search endpoint, API key, and index name.');
            return;
        } else {
            setDataSources(azureCogSearchEndpoint, azureCogSearchApiKey, azureCogSearchIndexName);
        }
    }

    if (!messageInitiated) {
        initMessages();
        messageInitiated = true;
    }

    document.getElementById('startSession').disabled = true;
    document.getElementById('configuration').hidden = true;

    try {
        const response = await fetch(`https://${privateEndpointEnabled ? privateEndpoint : `${cogSvcRegion}.tts.speech.microsoft.com`}/cognitiveservices/avatar/relay/token/v1`, {
            headers: { "Ocp-Apim-Subscription-Key": cogSvcSubKey }
        });
        const responseData = await response.json();
        const { Urls, Username, Password } = responseData;
        setupWebRTC(Urls[0], Username, Password);
    } catch (error) {
        console.error(`Failed to get relay token: ${error}`);
    }
}

// Disconnect from avatar service
function disconnectAvatar() {
    avatarSynthesizer?.close();
    speechRecognizer?.stopContinuousRecognitionAsync(() => speechRecognizer.close());
    sessionActive = false;
}

// Setup WebRTC
function setupWebRTC(iceServerUrl, iceServerUsername, iceServerCredential) {
    peerConnection = new RTCPeerConnection({
        iceServers: [{
            urls: [iceServerUrl],
            username: iceServerUsername,
            credential: iceServerCredential
        }]
    });

    peerConnection.ontrack = event => {
        const remoteVideoDiv = document.getElementById('remoteVideo');
        Array.from(remoteVideoDiv.childNodes).forEach(child => {
            if (child.localName === event.track.kind) {
                remoteVideoDiv.removeChild(child);
            }
        });

        if (event.track.kind === 'audio') {
            const audioElement = document.createElement('audio');
            audioElement.id = 'audioPlayer';
            audioElement.srcObject = event.streams[0];
            audioElement.autoplay = true;
            audioElement.onplaying = () => console.log(`WebRTC ${event.track.kind} channel connected.`);
            remoteVideoDiv.appendChild(audioElement);
        }

        if (event.track.kind === 'video') {
            document.getElementById('remoteVideo').style.width = '0.1px';
            if (!document.getElementById('useLocalVideoForIdle').checked) {
                document.getElementById('chatHistory').hidden = true;
            }

            const videoElement = document.createElement('video');
            videoElement.id = 'videoPlayer';
            videoElement.srcObject = event.streams[0];
            videoElement.autoplay = true;
            videoElement.playsInline = true;
            videoElement.onplaying = () => {
                console.log(`WebRTC ${event.track.kind} channel connected.`);
                document.getElementById('microphone').disabled = false;
                document.getElementById('stopSession').disabled = false;
                document.getElementById('remoteVideo').style.width = '960px';
                document.getElementById('chatHistory').hidden = false;
                document.getElementById('showTypeMessage').disabled = false;

                if (document.getElementById('useLocalVideoForIdle').checked) {
                    document.getElementById('localVideo').hidden = true;
                    lastSpeakTime = new Date();
                }

                setTimeout(() => { sessionActive = true }, 5000); // Set session active after 5 seconds
            };
            remoteVideoDiv.appendChild(videoElement);
        }
    };

    peerConnection.oniceconnectionstatechange = () => {
        console.log(`WebRTC status: ${peerConnection.iceConnectionState}`);
        if (peerConnection.iceConnectionState === 'disconnected' && document.getElementById('useLocalVideoForIdle').checked) {
            document.getElementById('localVideo').hidden = false;
            document.getElementById('remoteVideo').style.width = '0.1px';
        }
    };

    peerConnection.addTransceiver('video', { direction: 'sendrecv' });
    peerConnection.addTransceiver('audio', { direction: 'sendrecv' });

    avatarSynthesizer.startAvatarAsync(peerConnection).then(r => {
        if (r.reason === SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
            console.log(`[${new Date().toISOString()}] Avatar started. Result ID: ${r.resultId}`);
        } else {
            console.error(`[${new Date().toISOString()}] Unable to start avatar. Result ID: ${r.resultId}`);
            if (r.reason === SpeechSDK.ResultReason.Canceled) {
                const cancellationDetails = SpeechSDK.CancellationDetails.fromResult(r);
                if (cancellationDetails.reason === SpeechSDK.CancellationReason.Error) {
                    console.error(cancellationDetails.errorDetails);
                }
            }
            document.getElementById('startSession').disabled = false;
            document.getElementById('configuration').hidden = false;
        }
    }).catch(error => {
        console.error(`[${new Date().toISOString()}] Avatar failed to start. Error: ${error}`);
        document.getElementById('startSession').disabled = false;
        document.getElementById('configuration').hidden = false;
    });
}

// Initialize messages
function initMessages() {
    messages = [];
    if (dataSources.length === 0) {
        const systemPrompt = document.getElementById('prompt').value;
        messages.push({ role: 'system', content: systemPrompt });
    }
}

// Set data sources for chat API
function setDataSources(azureCogSearchEndpoint, azureCogSearchApiKey, azureCogSearchIndexName) {
    dataSources.push({
        type: 'AzureCognitiveSearch',
        parameters: {
            endpoint: azureCogSearchEndpoint,
            key: azureCogSearchApiKey,
            indexName: azureCogSearchIndexName,
            semanticConfiguration: '',
            queryType: 'simple',
            fieldsMapping: {
                contentFieldsSeparator: '\n',
                contentFields: ['content'],
                filepathField: null,
                titleField: 'title',
                urlField: null
            },
            inScope: true,
            roleInformation: document.getElementById('prompt').value
        }
    });
}

// Do HTML encoding on given text
function htmlEncode(text) {
    const entityMap = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
        '/': '&#x2F;'
    };

    return String(text).replace(/[&<>"'/]/g, match => entityMap[match]);
}

// Speak the given text
function speak(text, endingSilenceMs = 0) {
    if (isSpeaking) {
        spokenTextQueue.push(text);
        return;
    }

    speakNext(text, endingSilenceMs);
}

function speakNext(text, endingSilenceMs = 0) {
    const ttsVoice = document.getElementById('ttsVoice').value;
    const personalVoiceSpeakerProfileID = document.getElementById('personalVoiceSpeakerProfileID').value;
    let ssml = `<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='http://www.w3.org/2001/mstts' xml:lang='en-US'><voice name='${ttsVoice}'><mstts:ttsembedding speakerProfileId='${personalVoiceSpeakerProfileID}'><mstts:leadingsilence-exact value='0'/>${htmlEncode(text)}</mstts:ttsembedding></voice></speak>`;
    if (endingSilenceMs > 0) {
        ssml = `<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='http://www.w3.org/2001/mstts' xml:lang='en-US'><voice name='${ttsVoice}'><mstts:ttsembedding speakerProfileId='${personalVoiceSpeakerProfileID}'><mstts:leadingsilence-exact value='0'/>${htmlEncode(text)}<break time='${endingSilenceMs}ms' /></mstts:ttsembedding></voice></speak>`;
    }

    lastSpeakTime = new Date();
    isSpeaking = true;
    document.getElementById('stopSpeaking').disabled = false;
    avatarSynthesizer.speakSsmlAsync(ssml).then(result => {
        if (result.reason === SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
            console.log(`Speech synthesized to speaker for text [${text}]. Result ID: ${result.resultId}`);
            lastSpeakTime = new Date();
        } else {
            console.error(`Error occurred while speaking the SSML. Result ID: ${result.resultId}`);
        }

        if (spokenTextQueue.length > 0) {
            speakNext(spokenTextQueue.shift());
        } else {
            isSpeaking = false;
            document.getElementById('stopSpeaking').disabled = true;
        }
    }).catch(error => {
        console.error(`Error occurred while speaking the SSML: [${error}]`);

        if (spokenTextQueue.length > 0) {
            speakNext(spokenTextQueue.shift());
        } else {
            isSpeaking = false;
            document.getElementById('stopSpeaking').disabled = true;
        }
    });
}

function stopSpeaking() {
    spokenTextQueue = [];
    avatarSynthesizer.stopSpeakingAsync().then(() => {
        isSpeaking = false;
        document.getElementById('stopSpeaking').disabled = true;
        console.log(`[${new Date().toISOString()}] Stop speaking request sent.`);
    }).catch(error => {
        console.error(`Error occurred while stopping speaking: ${error}`);
    });
}

function handleUserQuery(userQuery) {
    messages.push({ role: 'user', content: userQuery });
    const chatHistoryTextArea = document.getElementById('chatHistory');
    if (chatHistoryTextArea.innerHTML !== '' && !chatHistoryTextArea.innerHTML.endsWith('\n\n')) {
        chatHistoryTextArea.innerHTML += '\n\n';
    }

    chatHistoryTextArea.innerHTML += `User: ${userQuery}\n\n`;
    chatHistoryTextArea.scrollTop = chatHistoryTextArea.scrollHeight;

    if (isSpeaking) {
        stopSpeaking();
    }

    if (dataSources.length > 0 && enableQuickReply) {
        speak(getQuickReply(), 2000);
    }

    const azureOpenAIEndpoint = document.getElementById('azureOpenAIEndpoint').value;
    const azureOpenAIApiKey = document.getElementById('azureOpenAIApiKey').value;
    const azureOpenAIDeploymentName = document.getElementById('azureOpenAIDeploymentName').value;

    let url = `${azureOpenAIEndpoint}/openai/deployments/${azureOpenAIDeploymentName}/chat/completions?api-version=2023-06-01-preview`;
    let body = JSON.stringify({
        messages: messages,
        stream: true
    });

    if (dataSources.length > 0) {
        url = `${azureOpenAIEndpoint}/openai/deployments/${azureOpenAIDeploymentName}/extensions/chat/completions?api-version=2023-06-01-preview`;
        body = JSON.stringify({
            dataSources: dataSources,
            messages: messages,
            stream: true
        });
    }

    let assistantReply = '';
    let toolContent = '';
    let spokenSentence = '';
    let displaySentence = '';

    fetch(url, {
        method: 'POST',
        headers: {
            'api-key': azureOpenAIApiKey,
            'Content-Type': 'application/json'
        },
        body: body
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`Chat API response status: ${response.status} ${response.statusText}`);
        }

        chatHistoryTextArea.innerHTML += 'Digital Lillian: ';

        const reader = response.body.getReader();

        // Function to recursively read chunks from the stream
        function read(previousChunkString = '') {
            return reader.read().then(({ value, done }) => {
                if (done) {
                    return;
                }

                let chunkString = new TextDecoder().decode(value, { stream: true });
                if (previousChunkString) {
                    chunkString = previousChunkString + chunkString;
                }

                if (!chunkString.endsWith('}\n\n') && !chunkString.endsWith('[DONE]\n\n')) {
                    return read(chunkString);
                }

                chunkString.split('\n\n').forEach(line => {
                    try {
                        if (line.startsWith('data:') && !line.endsWith('[DONE]')) {
                            const responseJson = JSON.parse(line.substring(5).trim());
                            let responseToken = undefined;
                            if (dataSources.length === 0) {
                                responseToken = responseJson.choices[0].delta.content;
                            } else {
                                const role = responseJson.choices[0].messages[0].delta.role;
                                if (role === 'tool') {
                                    toolContent = responseJson.choices[0].messages[0].delta.content;
                                } else {
                                    responseToken = responseJson.choices[0].messages[0].delta.content;
                                    if (responseToken && byodDocRegex.test(responseToken)) {
                                        responseToken = responseToken.replace(byodDocRegex, '').trim();
                                    }
                                    if (responseToken === '[DONE]') {
                                        responseToken = undefined;
                                    }
                                }
                            }

                            if (responseToken) {
                                assistantReply += responseToken;
                                displaySentence += responseToken;

                                if (responseToken === '\n' || responseToken === '\n\n') {
                                    speak(spokenSentence.trim());
                                    spokenSentence = '';
                                } else {
                                    responseToken = responseToken.replace(/\n/g, '');
                                    spokenSentence += responseToken;

                                    if (responseToken.length === 1 || responseToken.length === 2) {
                                        if (sentenceLevelPunctuations.includes(responseToken[0])) {
                                            speak(spokenSentence.trim());
                                            spokenSentence = '';
                                        }
                                    }
                                }
                            }
                        }
                    } catch (error) {
                        console.error(`Error occurred while parsing the response: ${error}`);
                        console.log(chunkString);
                    }
                });

                chatHistoryTextArea.innerHTML += displaySentence;
                chatHistoryTextArea.scrollTop = chatHistoryTextArea.scrollHeight;
                displaySentence = '';

                return read();
            });
        }

        return read();
    })
    .then(() => {
        if (spokenSentence) {
            speak(spokenSentence.trim());
            spokenSentence = '';
        }

        if (dataSources.length > 0) {
            messages.push({ role: 'tool', content: toolContent });
        }

        messages.push({ role: 'assistant', content: assistantReply });
    })
    .catch(error => {
        console.error(`Error occurred while fetching chat completions: ${error}`);
    });
}

function getQuickReply() {
    return quickReplies[Math.floor(Math.random() * quickReplies.length)];
}

function checkHung() {
    const videoElement = document.getElementById('videoPlayer');
    if (videoElement && sessionActive) {
        const videoTime = videoElement.currentTime;
        setTimeout(() => {
            if (videoElement.currentTime === videoTime && sessionActive) {
                sessionActive = false;
                if (document.getElementById('autoReconnectAvatar').checked) {
                    console.log(`[${new Date().toISOString()}] The video stream got disconnected, need reconnect.`);
                    connectAvatar();
                }
            }
        }, 5000);
    }
}

function checkLastSpeak() {
    if (!lastSpeakTime) return;
    const currentTime = new Date();
    if (currentTime - lastSpeakTime > 15000 && document.getElementById('useLocalVideoForIdle').checked && sessionActive && !isSpeaking) {
        disconnectAvatar();
        document.getElementById('localVideo').hidden = false;
        document.getElementById('remoteVideo').style.width = '0.1px';
        sessionActive = false;
    }
}

window.onload = () => {
    setInterval(() => {
        checkHung();
        checkLastSpeak();
    }, 5000);
};

window.startSession = () => {
    if (document.getElementById('useLocalVideoForIdle').checked) {
        document.getElementById('startSession').disabled = true;
        document.getElementById('configuration').hidden = true;
        document.getElementById('microphone').disabled = false;
        document.getElementById('stopSession').disabled = false;
        document.getElementById('localVideo').hidden = false;
        document.getElementById('remoteVideo').style.width = '0.1px';
        document.getElementById('chatHistory').hidden = false;
        document.getElementById('showTypeMessage').disabled = false;
        return;
    }

    connectAvatar();
};

window.stopSession = () => {
    document.getElementById('startSession').disabled = false;
    document.getElementById('microphone').disabled = true;
    document.getElementById('stopSession').disabled = true;
    document.getElementById('configuration').hidden = false;
    document.getElementById('chatHistory').hidden = true;
    document.getElementById('showTypeMessage').checked = false;
    document.getElementById('showTypeMessage').disabled = true;
    document.getElementById('userMessageBox').hidden = true;
    if (document.getElementById('useLocalVideoForIdle').checked) {
        document.getElementById('localVideo').hidden = true;
    }

    disconnectAvatar();
};

window.clearChatHistory = () => {
    document.getElementById('chatHistory').innerHTML = '';
    initMessages();
};

window.microphone = () => {
    const microphoneButton = document.getElementById('microphone');

    if (microphoneButton.innerHTML === 'Stop Microphone') {
        microphoneButton.disabled = true;
        speechRecognizer.stopContinuousRecognitionAsync(() => {
            microphoneButton.innerHTML = 'Start Microphone';
            microphoneButton.style.backgroundColor = '';
            microphoneButton.disabled = false;
        }, err => {
            console.error("Failed to stop continuous recognition:", err);
            microphoneButton.disabled = false;
        });
        return;
    }

    if (document.getElementById('useLocalVideoForIdle').checked && !sessionActive) {
        connectAvatar();
        setTimeout(() => {
            document.getElementById('audioPlayer').play();
        }, 5000);
    } else {
        document.getElementById('audioPlayer').play();
    }

    microphoneButton.disabled = true;
    speechRecognizer.recognized = async (s, e) => {
        if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
            const userQuery = e.result.text.trim();
            if (!userQuery) return;

            if (!document.getElementById('continuousConversation').checked) {
                microphoneButton.disabled = true;
                speechRecognizer.stopContinuousRecognitionAsync(() => {
                    microphoneButton.innerHTML = 'Start Microphone';
                    microphoneButton.style.backgroundColor = '';
                    microphoneButton.disabled = false;
                }, err => {
                    console.error("Failed to stop continuous recognition:", err);
                    microphoneButton.disabled = false;
                });
            }

            handleUserQuery(userQuery);
        }
    };

    speechRecognizer.startContinuousRecognitionAsync(() => {
        microphoneButton.innerHTML = 'Stop Microphone';
        microphoneButton.style.backgroundColor = 'green';
        microphoneButton.disabled = false;
    }, err => {
        console.error("Failed to start continuous recognition:", err);
        microphoneButton.disabled = false;
    });
};

window.updataEnableOyd = () => {
    document.getElementById('cogSearchConfig').hidden = !document.getElementById('enableOyd').checked;
};

window.updateTypeMessageBox = () => {
    const userMessageBox = document.getElementById('userMessageBox');
    if (document.getElementById('showTypeMessage').checked) {
        userMessageBox.hidden = false;
        userMessageBox.addEventListener('keyup', e => {
            if (e.key === 'Enter') {
                const userQuery = userMessageBox.value.trim();
                if (userQuery) {
                    handleUserQuery(userQuery);
                    userMessageBox.value = '';
                }
            }
        });
    } else {
        userMessageBox.hidden = true;
    }
};

window.updateLocalVideoForIdle = () => {
    document.getElementById('showTypeMessageCheckbox').hidden = document.getElementById('useLocalVideoForIdle').checked;
};

window.updatePrivateEndpoint = () => {
    document.getElementById('showPrivateEndpointCheckBox').hidden = !document.getElementById('enablePrivateEndpoint').checked;
};

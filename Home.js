Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        // Office is ready, initialize 
        document.getElementById("checkGrammarButton").onclick = checkGrammar;
        document.getElementById("paraphraseButton").onclick = paraphraseSelectedText;

       
        document.getElementById("addIgnorePatternButton").onclick = addIgnorePattern;
        document.getElementById("addSpecialTermButton").onclick = addSpecialTerm;
        document.getElementById("suggestionButton").onclick = getSuggestionsForSelectedText;

        document.getElementById("analyzeStructureButton").onclick = () => {
            analyzeDocumentStructure();
        };

        // Load settings and update UI when the add-in is loaded
        loadSettingsAndUpdateUI();



    }
});




// This function is triggered when the "suggestionButton" is clicked
function getSuggestionsForSelectedText() {
    Word.run(async context => {
        const range = context.document.getSelection();
        range.load('text');
        await context.sync();

        const selectedText = range.text.trim();
        if (selectedText) {
            const apiUrl = 'https://api.languagetool.org/v2/check';
            const options = {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: `language=en-US&text=${encodeURIComponent(selectedText)}`
            };

            try {
                const response = await fetch(apiUrl, options);
                const data = await response.json();

                // Process and display suggestions
                if (data.matches && data.matches.length > 0) {
                    displaySuggestions(data.matches.map(match => {
                        return {
                            message: match.message,
                            suggestion: match.replacements.length > 0 ? match.replacements[0].value : 'No suggestion'
                        };
                    }));
                }
            } catch (error) {
                console.error('Error:', error);
                displayMessage('Error fetching suggestions: ' + error.message);
            }
        }
    }).catch(error => {
        console.error('Error:', error);
        displayMessage('Error: ' + error);
    });
}

function displaySuggestions(suggestions) {
    const suggestionsContainer = document.getElementById('suggestionsContainer');
    suggestionsContainer.innerHTML = '';
    suggestions.forEach(suggestion => {
        const suggestionElement = document.createElement('div');
        suggestionElement.innerHTML = `<strong>Issue:</strong> ${suggestion.message} <br>
                                       <strong>Suggestion:</strong> ${suggestion.suggestion}`;
        suggestionsContainer.appendChild(suggestionElement);
    });
}



// Implement the displayMessage function to show messages 
function displayMessage(message) {
    const messageContainer = document.getElementById('messageContainer');
    messageContainer.textContent = message;
}

async function analyzeDocumentStructure() {
    await Word.run(async context => {
        const body = context.document.body;
        context.load(body, 'text');
        await context.sync();

        const formdata = new FormData();
        formdata.append("key", "352923795ffd7bfa85dbaa30879ce27a");
        formdata.append("txt", body.text);

        const requestOptions = {
            method: 'POST',
            body: formdata,
            redirect: 'follow'
        };

        try {
            const response = await fetch("https://api.meaningcloud.com/documentstructure-1.0", requestOptions);
            const result = await response.json();
            document.getElementById("messageContainer").textContent = formatApiResponse(result);
        } catch (error) {
            console.error('Error:', error);
            document.getElementById("messageContainer").textContent = 'Error: ' + error.message;
        }
    });
}

function formatApiResponse(apiResponse) {
    let formattedText = '';

    // Check for headings and append them to the formatted text
    if (apiResponse.heading_list && apiResponse.heading_list.length > 0) {
        formattedText += 'Headings:\n';
        apiResponse.heading_list.forEach(heading => {
            formattedText += `- ${heading.text}\n`;
        });
    } else {
        formattedText += 'No headings found.\n';
    }

    // Check for other elements like abstract, title, etc. and append them to the formatted text



    if (apiResponse.title) {
        formattedText += `Title: ${apiResponse.title}\n`;
    }

    // Example for abstract
    if (apiResponse.abstract_list && apiResponse.abstract_list.length > 0) {
        formattedText += `Abstract: ${apiResponse.abstract_list.join(' ')}\n`;
    }

    // Example for email info
    if (apiResponse.emails_info && apiResponse.emails_info.from) {
        formattedText += `Email from: ${apiResponse.emails_info.from}\n`;
    }

    // Return the formatted text
    return formattedText;
}


function checkGrammar() {
    // First, we need to retrieve the ignored patterns and special terms from the settings.
    Word.run(function (context) {
        var body = context.document.body;
        var range = body.getRange();
        range.load("text");

        return context.sync().then(function () {
            var textToCheck = range.text;
            var ignoredPatterns = getIgnoredPatternsFromSettings();
            var specialTerms = getSpecialTermsFromSettings();

            // Now we apply the ignored patterns to the text
            ignoredPatterns.forEach(pattern => {
                const regex = new RegExp(pattern, 'gi');
                console.log(`Applying ignore pattern: ${pattern}`);
                textToCheck = textToCheck.replace(regex, '');
                console.log(`Text after applying ignore pattern: ${textToCheck}`);
            });

            // Then we send the text to the LanguageTool API
            var requestBody = new URLSearchParams();
            requestBody.append("language", "en-US");
            requestBody.append("text", textToCheck);

            fetch("https://api.languagetool.org/v2/check", {
                method: "POST",
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded",
                },
                body: requestBody,
            })
                .then(response => response.json())
                .then(data => {


                    // Apply the logic to filter out the special terms from the errors
                    var errors = data.matches.filter(error => {
                        var errorText = error.context.text;
                        var isSpecialTerm = specialTerms.some(term => errorText.toLowerCase().includes(term.toLowerCase()));

                        return !isSpecialTerm;
                    });


                    displayErrors(errors);
                })
                .catch(error => {
                    console.error("Error checking grammar:", error);
                });
        });
    }).catch(function (error) {
        console.error("Error with Word JavaScript API:", error);
    });
}

// Function for paraphrasing text


async function paraphraseSelectedText() {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");

            await context.sync();

            const selectedText = range.text;

            if (selectedText) {
                const url = 'https://rewriter-paraphraser-text-changer-multi-language.p.rapidapi.com/rewrite';
                const apiKey = '76efc77793msh95945de9527f1aap16c049jsn0b93303c3b07'; // Replace with your API key

                const options = {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-RapidAPI-Key': apiKey,
                        'X-RapidAPI-Host': 'rewriter-paraphraser-text-changer-multi-language.p.rapidapi.com'
                    },
                    body: JSON.stringify({
                        language: 'en',
                        strength: 3,
                        text: selectedText
                    })
                };

                const response = await fetch(url, options);

                if (!response.ok) {
                    throw new Error('Network response was not ok.');
                }

                const result = await response.json();

                if (result && result.rewrite) {
                    const paraphrasedText = result.rewrite;
                    range.insertText(paraphrasedText, Word.InsertLocation.replace);
                    console.log('Text paraphrased successfully!');
                } else {
                    console.log('Error: Paraphrased text not found in the API response.');
                }
            } else {
                console.log('No text selected for paraphrasing.');
            }
        });
    } catch (error) {
        console.error('Error paraphrasing text:', error);
    }
}




function addIgnorePattern() {
    var pattern = document.getElementById("ignorePatternInput").value;
    // Add pattern to settings and update UI
    updateSettingsAndUI('ignoredPatterns', pattern, "ignorePatternList");
}

// Function to add a special term
function addSpecialTerm() {
    var term = document.getElementById("specialTermInput").value;
    // Add term to settings and update UI
    updateSettingsAndUI('specialTerms', term, "specialTermList");
}

function getIgnoredPatternsFromSettings() {
    // Assuming the patterns are saved as a stringified JSON array in the settings
    var ignoredPatternsJson = Office.context.document.settings.get('ignoredPatterns');
    return ignoredPatternsJson ? JSON.parse(ignoredPatternsJson) : [];
}

function getSpecialTermsFromSettings() {
    // Assuming the terms are saved as a stringified JSON array in the settings
    var specialTermsJson = Office.context.document.settings.get('specialTerms');
    return specialTermsJson ? JSON.parse(specialTermsJson) : [];
}


// Update settings and UI helper function
function updateSettingsAndUI(settingKey, value, listElementId) {
    if (!value) return; // Do nothing if the value is empty

    Word.run(function (context) {
        var settings = Office.context.document.settings;
        var items = settings.get(settingKey) ? JSON.parse(settings.get(settingKey)) : [];
        items.push(value);
        // Correctly stringify the array before saving it
        settings.set(settingKey, JSON.stringify(items));
        return context.sync().then(function () {
            settings.saveAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    displayMessage("Error: " + asyncResult.error.message);
                } else {
                    // Update the UI
                    addToListUI(value, listElementId);
                    displayMessage("Added new item to " + settingKey);
                }
            });
        });
    }).catch(function (error) {
        displayMessage("Error: " + error);
    });
}


// Load settings and update UI
function loadSettingsAndUpdateUI() {
    Word.run(function (context) {
        var settings = Office.context.document.settings;
        // Request to load the 'ignoredPatterns' and 'specialTerms' settings
        settings.load(['ignoredPatterns', 'specialTerms']);
        return context.sync().then(function () {
            // Parse the settings to get the arrays
            var ignoredPatterns = settings.get('ignoredPatterns') ? JSON.parse(settings.get('ignoredPatterns')) : [];
            var specialTerms = settings.get('specialTerms') ? JSON.parse(settings.get('specialTerms')) : [];

            // Update the UI with the loaded settings
            updateIgnoredPatternsUI(ignoredPatterns);
            updateSpecialTermsUI(specialTerms);
        });
    }).catch(function (error) {
        displayMessage("Error: " + error);
    });
}


// Function to add items to the list UI
function addToListUI(value, listElementId) {
    var listElement = document.getElementById(listElementId);
    var listItem = document.createElement("li");
    listItem.textContent = value;
    listElement.appendChild(listItem);
}

// Function to display status messages
function displayMessage(message) {
    var messageContainer = document.getElementById("messageContainer");
    messageContainer.textContent = message;
    setTimeout(() => { messageContainer.textContent = ''; }, 5000); // Clear message after 5 seconds
}

function displayErrors(errors) {
    // You can display the errors in a specific element in your HTML.
    var errorElement = document.getElementById("errorList");

    // Clear any previous error messages
    errorElement.innerHTML = "";

    if (errors && errors.length > 0) {
        // Loop through the errors and display each one
        for (var i = 0; i < errors.length; i++) {
            var error = errors[i];
            var errorMessage = error.message;
            var errorContext = error.context;

            // Create a list item to display the error message and context
            var listItem = document.createElement("li");
            listItem.textContent = errorMessage;

            if (errorContext) {
                if (Array.isArray(errorContext)) {
                    // If errorContext is an array, display it
                    listItem.textContent += " (Context: " + errorContext.map(function (ctx) {
                        return ctx.text;
                    }).join(" ") + ")";
                } else if (typeof errorContext === "string") {
                    // If errorContext is a string, simply append it
                    listItem.textContent += " (Context: " + errorContext + ")";
                }
            }

            // Append the list item to the error element
            errorElement.appendChild(listItem);
        }
    } else {
        // If no errors are found, display a message indicating that
        errorElement.textContent = "No grammar or spelling errors found.";
    }
}


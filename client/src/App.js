// Importing necessary modules and assets
import React, { useState } from 'react'; // React and useState hook for functional component
import logo from './logo.svg'; // Logo image for the app
import './App.css'; // CSS file for styling

// Defining the App component
const App = () => {
  // State variables
  const [folder_url, setFolderUrl] = useState(''); // Variable for storing the folder ID
  const [gradebook_url, setGradebookUrl] = useState(''); // Variable for storing the gradebook ID
  const [message, setMessage] = useState(''); // Variable for storing messages to display to the user
  const [isPressed, setIsPressed] = useState(false);

  // Function to handle form submission
  const handleSubmit = async (event) => {
    setIsPressed(true);
    setMessage('Processing, please wait...');
    event.preventDefault(); // Preventing default form submission behavior
    // Creating query parameters for the API request
    const queryParams = new URLSearchParams({ folder_url, gradebook_url }).toString();
    // Constructing the request URL with query parameters
    const requestUrl = `http://173.8.188.76:5000//run_gradebook?${queryParams}`;

    try {
      // Making an asynchronous request to the server
      const response = await fetch(requestUrl);
      const data = await response.json(); // Parsing the response data as JSON
      setMessage(data.message); // Setting the message state with the response message
    } catch (error) {
      // Error handling
      console.error('Error:', error);
      setMessage('Failed to process request'); // Setting error message
    }
  };

  const buttonStyle = isPressed ? { backgroundColor: 'green', color: 'white' } : {};

  // HTML: Render method for the component
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <h1>
          Gradebook Jana-Rator
        </h1>

        <form onSubmit={handleSubmit}>
          <div>
            <label>
              Google Drive Rubrics Folder ID:
              <input
                type="text"
                className="long-input"
                value={folder_url}
                onChange={(e) => setFolderUrl(e.target.value)}
              />
            </label>
          </div>

          <div>
            <label>
              Google Sheets Gradebook ID:
              <input
                type="text"
                className="long-input"
                value={gradebook_url}
                onChange={(e) => setGradebookUrl(e.target.value)}
              />
            </label>
          </div>

          <button style={buttonStyle} type="submit">Run Gradebook</button>
        </form>
        {message && <p>{message}</p>}
      </header>
    </div>
  );
};

// Exporting the App component for use in other parts of the application
export default App;

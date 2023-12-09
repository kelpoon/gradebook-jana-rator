import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import App from './App';

// This line mocks the global fetch function used for making API requests.
// It simulates a successful response with the message "Success".
global.fetch = jest.fn(() =>
  Promise.resolve({
    json: () => Promise.resolve({ message: "Success" })
  })
);

// beforeEach is a setup function that runs before each test case.
// Here, it's clearing the mock data of fetch to ensure test isolation.
beforeEach(() => {
  fetch.mockClear();
});

// This block describes a series of test cases for the App component.
describe('App Component', () => {
  // Testing if the App component renders correctly.
  test('renders App component', () => {
    render(<App />);
    expect(screen.getByText(/Gradebook Jana-Rator/i)).toBeInTheDocument();
  });

  // Testing if the App component updates its state when input changes.
  test('updates state on input change', () => {
    render(<App />);
    const folderInput = screen.getByLabelText(/Google Drive Rubrics Folder ID:/i);
    fireEvent.change(folderInput, { target: { value: '123' } });
    expect(folderInput.value).toBe('123');
  });

  // Testing the form submission behavior.
  test('handles form submission', async () => {
    render(<App />);
    const button = screen.getByText(/Run Gradebook/i);
    fireEvent.click(button);

    // Waits for the fetch call to happen after form submission.
    await waitFor(() => {
      expect(fetch).toHaveBeenCalled();
    });
  });

  // Testing if a message is displayed after a successful fetch request.
  test('displays message after successful fetch', async () => {
    render(<App />);
    const button = screen.getByText(/Run Gradebook/i);
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("Success")).toBeInTheDocument();
    });
  });

  // Testing how the App component handles fetch failures.
  test('handles fetch failure', async () => {
    // Mocks a failed fetch request for this specific test.
    fetch.mockImplementationOnce(() => Promise.reject(new Error("Failed to fetch")));

    render(<App />);
    const button = screen.getByText(/Run Gradebook/i);
    fireEvent.click(button);

    // Checks if an error message is displayed after fetch failure.
    await waitFor(() => {
      expect(screen.getByText(/Failed to process request/i)).toBeInTheDocument();
    });
  });
});

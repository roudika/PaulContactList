// MSAL Configuration
const msalConfig = {
  auth: {
    clientId: "b1f8ddfa-6663-4192-9137-5c30eb6673ae",
    authority: "https://login.microsoftonline.com/2b21e8b5-c462-4f9d-952f-f47b9456b623",
    redirectUri: "https://roudika.github.io/PaulContactList/Index.html"
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  },
  system: {
    allowNativeBroker: false,
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case msal.LogLevel.Error:
            console.error(message);
            break;
          case msal.LogLevel.Info:
            console.info(message);
            break;
          case msal.LogLevel.Verbose:
            console.debug(message);
            break;
          case msal.LogLevel.Warning:
            console.warn(message);
            break;
        }
      }
    }
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Initialize application
document.addEventListener('DOMContentLoaded', () => {
  console.log("DOM Content Loaded");
  
  // Set up event listeners
  setupEventListeners();
  
  // Check if we're already authenticated
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts && currentAccounts.length > 0) {
    // Redirect to main page if already authenticated
    window.location.href = 'Index.html';
  }
});

function setupEventListeners() {
  // Add sign in button handler
  const signInButton = document.getElementById('signin');
  console.log("Setting up sign in button handler...");
  if (signInButton) {
    signInButton.addEventListener('click', (e) => {
      console.log("Sign in button clicked!");
      e.preventDefault();
      signIn();
    });
  } else {
    console.error("Sign in button not found!");
  }
}

async function signIn() {
  try {
    console.log("Starting sign in process...");
    const loginRequest = {
      scopes: ["User.Read", "GroupMember.Read.All"],
      prompt: "select_account"
    };

    console.log("Attempting redirect login...");
    await msalInstance.loginRedirect(loginRequest);
  } catch (error) {
    console.error("Error during sign in:", error);
    alert("Sign in failed. Please try again.");
  }
} 
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
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
          default:
            return;
        }
      }
    }
  }
};

// Global variables
let accessToken = "";
let allMembers = [];
let activeDepartment = null;
let currentSort = "name";

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Initialize application
document.addEventListener('DOMContentLoaded', () => {
  console.log("DOM Content Loaded");
  
  // Handle redirect promise
  msalInstance.handleRedirectPromise().then(response => {
    if (response) {
      // User is already signed in
      console.log("Got response from redirect:", response);
      msalInstance.setActiveAccount(response.account);
      showWelcomeUI(response.account);
      getTokenAndLoadMembers();
    } else {
      // Check if user is already signed in
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        console.log("Found existing account:", accounts[0]);
        msalInstance.setActiveAccount(accounts[0]);
        showWelcomeUI(accounts[0]);
        getTokenAndLoadMembers();
      } else {
        // No account found, redirect to login
        window.location.href = "login.html";
      }
    }
  }).catch(error => {
    console.error("Error handling redirect:", error);
    window.location.href = "login.html";
  });

  // Set up event listeners
  setupEventListeners();
});

function setupEventListeners() {
  // Add sign in button handler
  const signInButton = document.getElementById('signin');
  console.log("Setting up sign in button handler...");
  if (signInButton) {
    signInButton.addEventListener('click', (e) => {
      console.log("Sign in button clicked!");
      e.preventDefault(); // Prevent any default form submission
      signIn();
    });
  } else {
    console.error("Sign in button not found!");
  }

  // Keyboard shortcuts
  document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + K for search focus
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      document.getElementById('searchInput').focus();
    }
    // Esc to clear search and filters
    if (e.key === 'Escape') {
      document.getElementById('searchInput').value = '';
      document.getElementById('departmentSelect').value = '';
      activeDepartment = null;
      filterContacts();
    }
  });

  // Sort select change handler
  document.getElementById('sortSelect').addEventListener('change', (e) => {
    currentSort = e.target.value;
    filterContacts();
  });

  // Search input handler
  document.getElementById("searchInput").addEventListener("input", (e) => {
    const term = e.target.value.toLowerCase();
    const filtered = allMembers.filter(user =>
      (user.displayName || "").toLowerCase().includes(term) ||
      (user.mail || "").toLowerCase().includes(term) ||
      (user.department || "").toLowerCase().includes(term)
    );
    renderContactList(filtered);
  });

  // Dark mode toggle
  document.getElementById('darkModeToggle').addEventListener('click', () => {
    const html = document.documentElement;
    const isDark = html.getAttribute('data-bs-theme') === 'dark';
    html.setAttribute('data-bs-theme', isDark ? 'light' : 'dark');
    document.getElementById('darkModeToggle').innerHTML = 
      `<i class="bi bi-${isDark ? 'moon-fill' : 'sun-fill'}"></i>`;
  });

  // Logout handler
  document.getElementById('logoutBtn').addEventListener('click', () => {
    msalInstance.logoutPopup().then(() => {
      accessToken = "";
      allMembers = [];
      document.getElementById('userGreeting').classList.add('d-none');
      document.getElementById('logoutBtn').classList.add('d-none');
      document.getElementById('contactList').innerHTML = '';
      document.getElementById('totalContacts').textContent = '0';
      const signInModal = new bootstrap.Modal(document.getElementById('signInModal'));
      signInModal.show();
    });
  });
}

function closeSignInModal() {
  const signInModal = document.getElementById('signInModal');
  if (signInModal) {
    signInModal.style.display = 'none';
    signInModal.classList.remove('show');
    const backdrop = document.querySelector('.modal-backdrop');
    if (backdrop) backdrop.remove();
    document.body.classList.remove('modal-open');
    document.body.style.overflow = '';
    document.body.style.paddingRight = '';
  }
}

function showWelcomeUI(account) {
  if (!account) return;
  
  const userGreeting = document.getElementById('userGreeting');
  userGreeting.innerHTML = `<i class="bi bi-person-circle me-1"></i>Hello, ${account.name}!`;
  userGreeting.classList.remove('d-none');
  
  // Show logout button
  document.getElementById('logoutBtn').classList.remove('d-none');
  
  // Close the modal
  closeSignInModal();
  
  // Load contacts
  loadContacts();
}

async function getTokenAndLoadMembers() {
  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["User.Read", "GroupMember.Read.All"],
      account: msalInstance.getActiveAccount()
    });
    accessToken = tokenResponse.accessToken;
    await loadGroupMembers();
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      try {
        // Try popup first
        try {
          const tokenResponse = await msalInstance.acquireTokenPopup({
            scopes: ["User.Read", "GroupMember.Read.All"]
          });
          accessToken = tokenResponse.accessToken;
          await loadGroupMembers();
        } catch (popupError) {
          console.log("Token popup failed, falling back to redirect:", popupError);
          // Fallback to redirect
          await msalInstance.acquireTokenRedirect({
            scopes: ["User.Read", "GroupMember.Read.All"]
          });
        }
      } catch (err) {
        console.error(err);
        alert("Failed to acquire token. Please try signing in again.");
      }
    }
  }
}

function sortContacts(contacts) {
  switch (currentSort) {
    case "department":
      return contacts.sort((a, b) => {
        const deptA = (a.department || 'ZZZ').toLowerCase();
        const deptB = (b.department || 'ZZZ').toLowerCase();
        return deptA === deptB ? 
          a.displayName.localeCompare(b.displayName) : 
          deptA.localeCompare(deptB);
      });
    case "title":
      return contacts.sort((a, b) => {
        const titleA = (a.jobTitle || 'ZZZ').toLowerCase();
        const titleB = (b.jobTitle || 'ZZZ').toLowerCase();
        return titleA === titleB ? 
          a.displayName.localeCompare(b.displayName) : 
          titleA.localeCompare(titleB);
      });
    default:
      return contacts.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }
}

function filterContacts() {
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const filtered = allMembers.filter(user => {
    const matchesSearch = (user.displayName || '').toLowerCase().includes(searchTerm) ||
                        (user.mail || '').toLowerCase().includes(searchTerm);
    const matchesDepartment = !activeDepartment || user.department === activeDepartment;
    return matchesSearch && matchesDepartment;
  });
  renderContactList(sortContacts(filtered));
}

function renderDepartmentSelect(members) {
  const departments = [...new Set(members.map(m => m.department).filter(Boolean))].sort();
  const select = document.getElementById('departmentSelect');
  select.innerHTML = '<option value="">All Departments</option>';
  
  departments.forEach(dept => {
    const option = document.createElement('option');
    option.value = dept;
    option.textContent = dept;
    if (dept === activeDepartment) {
      option.selected = true;
    }
    select.appendChild(option);
  });
}

async function loadGroupMembers() {
  const groupId = "2ac0dfde-a4db-4e8a-af91-7fa805271a37";
  let endpoint = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=displayName,mail,mobilePhone,department,jobTitle,userPrincipalName,id&$top=999`;
  let allMembersArray = [];

  document.getElementById("loading").classList.remove("d-none");
  document.getElementById("skeletonList").classList.remove("d-none");
  document.getElementById("contactList").innerHTML = '';
  document.getElementById("loading").textContent = "Loading members...";

  try {
    while (endpoint) {
      const response = await fetch(endpoint, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });

      if (!response.ok) {
        throw new Error(`
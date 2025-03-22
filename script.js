// MSAL Configuration
const msalConfig = {
  auth: {
    clientId: "b1f8ddfa-6663-4192-9137-5c30eb6673ae",
    authority: "https://login.microsoftonline.com/2b21e8b5-c462-4f9d-952f-f47b9456b623",
    redirectUri: window.location.origin + "/PaulContactList/Index.html"
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
  console.log("Current URL:", window.location.href);
  console.log("MSAL Config:", msalConfig);
  
  // Handle redirect promise
  msalInstance.handleRedirectPromise().then(response => {
    console.log("Redirect promise response:", response);
    if (response) {
      // User is already signed in
      console.log("Got response from redirect:", response);
      msalInstance.setActiveAccount(response.account);
      showWelcomeUI(response.account);
      getTokenAndLoadMembers();
    } else {
      // Check if user is already signed in
      const accounts = msalInstance.getAllAccounts();
      console.log("Existing accounts:", accounts);
      if (accounts.length > 0) {
        console.log("Found existing account:", accounts[0]);
        msalInstance.setActiveAccount(accounts[0]);
        showWelcomeUI(accounts[0]);
        getTokenAndLoadMembers();
      }
    }
  }).catch(error => {
    console.error("Error handling redirect:", error);
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
      e.preventDefault();
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
    console.log("Logout button clicked");
    msalInstance.logoutRedirect().then(() => {
      console.log("Logout successful");
      accessToken = "";
      allMembers = [];
      document.getElementById('userGreeting').classList.add('d-none');
      document.getElementById('logoutBtn').classList.add('d-none');
      document.getElementById('signin').classList.remove('d-none');
      document.getElementById('contactList').innerHTML = '';
      document.getElementById('totalContacts').textContent = '0';
    }).catch(error => {
      console.error("Logout failed:", error);
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
  
  console.log("Showing welcome UI for account:", account);
  const userGreeting = document.getElementById('userGreeting');
  userGreeting.innerHTML = `<i class="bi bi-person-circle me-1"></i>Hello, ${account.name}!`;
  userGreeting.classList.remove('d-none');
  
  // Show logout button and hide sign in button
  document.getElementById('logoutBtn').classList.remove('d-none');
  document.getElementById('signin').classList.add('d-none');
  
  // Load contacts
  loadContacts();
}

function signIn() {
  console.log("Initiating sign in with redirect...");
  console.log("Current URL:", window.location.href);
  console.log("Redirect URI:", msalConfig.auth.redirectUri);
  
  msalInstance.loginRedirect({
    scopes: ["User.Read", "GroupMember.Read.All"]
  }).catch(error => {
    console.error("Sign in failed:", error);
    alert("Failed to sign in. Please try again.");
  });
}

async function getTokenAndLoadMembers() {
  try {
    console.log("Attempting to get token silently...");
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["User.Read", "GroupMember.Read.All"],
      account: msalInstance.getActiveAccount()
    });
    console.log("Got token successfully");
    accessToken = tokenResponse.accessToken;
    await loadGroupMembers();
  } catch (error) {
    console.error("Error getting token:", error);
    if (error instanceof msal.InteractionRequiredAuthError) {
      try {
        console.log("Token expired, trying redirect...");
        msalInstance.acquireTokenRedirect({
          scopes: ["User.Read", "GroupMember.Read.All"]
        });
      } catch (redirectError) {
        console.error("Redirect failed:", redirectError);
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

  document.getElementById("loading").classList.remove("d-none");
  document.getElementById("skeletonList").classList.remove("d-none");
  document.getElementById("contactList").innerHTML = '';
  document.getElementById("loading").textContent = "Loading members...";

  try {
    console.log("Fetching members from Graph API...");
    const response = await fetch(endpoint, {
      headers: { 
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Graph API error:", response.status, errorText);
      throw new Error(`Failed to fetch members: ${response.status} ${errorText}`);
    }

    const data = await response.json();
    console.log("Got members data:", data);
    
    allMembers = data.value;
    document.getElementById("totalContacts").textContent = allMembers.length;
    document.getElementById("totalContactsLarge").textContent = allMembers.length;
    
    renderDepartmentSelect(allMembers);
    renderContactList(sortContacts(allMembers));
  } catch (error) {
    console.error("Error loading members:", error);
    document.getElementById("loading").textContent = "Error loading members. Please try again.";
  } finally {
    document.getElementById("loading").classList.add("d-none");
    document.getElementById("skeletonList").classList.add("d-none");
  }
}

function renderContactList(contacts) {
  const contactList = document.getElementById('contactList');
  contactList.innerHTML = '';

  contacts.forEach(contact => {
    const card = document.createElement('div');
    card.className = 'col-12 col-md-6 col-lg-4';
    card.innerHTML = `
      <div class="card shadow-sm p-3 contact-card">
        <div class="d-flex align-items-center">
          <img src="https://graph.microsoft.com/v1.0/users/${contact.id}/photo/$value" 
               class="profile-pic me-3" 
               onerror="this.src='https://via.placeholder.com/48'"
               alt="${contact.displayName}">
          <div class="flex-grow-1">
            <h5 class="mb-1">${contact.displayName}</h5>
            <div class="card-text">
              <strong>Email:</strong> 
              <span class="copyable" data-copy="${contact.mail}">${contact.mail}</span>
            </div>
            ${contact.department ? `
              <div class="card-text">
                <strong>Dept:</strong> 
                <span class="badge badge-department">${contact.department}</span>
              </div>
            ` : ''}
          </div>
        </div>
      </div>
    `;
    contactList.appendChild(card);
  });
}

function loadContacts() {
  getTokenAndLoadMembers();
}
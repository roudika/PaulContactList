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
          case msal.LogLevel.Error:
            console.error(message);
            return;
          case msal.LogLevel.Info:
            console.info(message);
            return;
          case msal.LogLevel.Verbose:
            console.debug(message);
            return;
          case msal.LogLevel.Warning:
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

// Add profile picture cache
const profilePicCache = new Map();

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

async function showWelcomeUI(account) {
  if (!account) return;
  
  console.log("Showing welcome UI for account:", account);
  const userGreeting = document.getElementById('userGreeting');
  userGreeting.textContent = `Hello, ${account.name}!`;
  userGreeting.classList.remove('d-none');
  
  // Show logout button and hide sign in button
  document.getElementById('logoutBtn').classList.remove('d-none');
  document.getElementById('signin').classList.add('d-none');

  // Get user's ID from Graph API
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });

    if (response.ok) {
      const userData = await response.json();
      console.log("Got user data:", userData);
      
      // Load user's profile picture
      const userProfilePic = document.getElementById('userProfilePic');
      if (userProfilePic) {
        loadProfilePicture(userData.id, userProfilePic);
      }
    }
  } catch (error) {
    console.error("Failed to get user data:", error);
  }
  
  // Load contacts
  loadContacts();
}

function signIn() {
  console.log("Initiating sign in with redirect...");
  console.log("Current URL:", window.location.href);
  console.log("Redirect URI:", msalConfig.auth.redirectUri);
  
  // Use only redirect method
  msalInstance.loginRedirect({
    scopes: ["User.Read", "GroupMember.Read.All"]
  }).catch(error => {
    console.error("Sign in failed:", error);
    alert("Failed to sign in. Please try again. Error: " + error.message);
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
    
    // Create a unique ID for the image
    const imgId = `profile-pic-${contact.id}`;
    
    card.innerHTML = `
      <div class="card shadow-sm p-3 contact-card" data-contact-id="${contact.id}">
        <div class="d-flex align-items-start">
          <div class="profile-pic-container me-3">
            <img id="${imgId}"
                 class="profile-pic" 
                 src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='36' height='36' viewBox='0 0 36 36'%3E%3Crect width='36' height='36' fill='%23e9ecef'/%3E%3Cpath d='M18 18c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z' fill='%23999'/%3E%3C/svg%3E"
                 alt="${contact.displayName}"
                 data-user-id="${contact.id}">
          </div>
          <div class="flex-grow-1">
            <h5 class="mb-1">${contact.displayName}</h5>
            ${contact.jobTitle ? `
              <div class="badge badge-title mb-2">${contact.jobTitle}</div>
            ` : ''}
            <div class="card-text">
              <i class="bi bi-envelope me-1"></i>
              <span class="copyable" data-copy="${contact.mail}">${contact.mail}</span>
            </div>
            ${contact.mobilePhone ? `
              <div class="card-text">
                <i class="bi bi-telephone me-1"></i>
                <span class="copyable" data-copy="${contact.mobilePhone}">${contact.mobilePhone}</span>
              </div>
            ` : ''}
          </div>
        </div>
      </div>
    `;
    contactList.appendChild(card);

    // Add click handler to the card
    const cardElement = card.querySelector('.contact-card');
    cardElement.addEventListener('click', () => showContactDetails(contact));

    // Load profile picture if not in cache
    const img = document.getElementById(imgId);
    if (img && !profilePicCache.has(contact.id)) {
      loadProfilePicture(contact.id, img);
    } else if (img && profilePicCache.has(contact.id)) {
      img.src = profilePicCache.get(contact.id);
    }
  });
}

function showContactDetails(contact) {
  const modalBody = document.getElementById('modalBodyContent');
  modalBody.innerHTML = `
    <div class="d-flex align-items-center mb-4">
      <div class="profile-pic-container me-3" style="width: 96px; height: 96px;">
        <img id="modal-profile-pic"
             class="profile-pic" 
             src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='96' height='96' viewBox='0 0 96 96'%3E%3Crect width='96' height='96' fill='%23e9ecef'/%3E%3Cpath d='M48 48c4.42 0 8-3.58 8-8s-3.58-8-8-8-8 3.58-8 8 3.58 8 8 8zm0 4c-5.34 0-16 2.68-16 8v4h32v-4c0-5.32-10.66-8-16-8z' fill='%23999'/%3E%3C/svg%3E"
             alt="${contact.displayName}"
             style="width: 96px; height: 96px;">
      </div>
      <div>
        <h4 class="mb-1">${contact.displayName}</h4>
        ${contact.jobTitle ? `<div class="badge badge-title mb-2">${contact.jobTitle}</div>` : ''}
        ${contact.department ? `<div class="text-muted mb-1"><i class="bi bi-building me-1"></i>${contact.department}</div>` : ''}
      </div>
    </div>
    <div class="contact-details">
      <div class="detail-item">
        <i class="bi bi-envelope me-2"></i>
        <span class="copyable" data-copy="${contact.mail}">${contact.mail}</span>
      </div>
      ${contact.mobilePhone ? `
        <div class="detail-item">
          <i class="bi bi-telephone me-2"></i>
          <span class="copyable" data-copy="${contact.mobilePhone}">${contact.mobilePhone}</span>
        </div>
      ` : ''}
      ${contact.userPrincipalName ? `
        <div class="detail-item">
          <i class="bi bi-person me-2"></i>
          <span>${contact.userPrincipalName}</span>
        </div>
      ` : ''}
    </div>
  `;

  // Load profile picture for modal
  const modalImg = document.getElementById('modal-profile-pic');
  if (modalImg) {
    loadProfilePicture(contact.id, modalImg);
  }

  // Show the modal
  const modal = new bootstrap.Modal(document.getElementById('detailsModal'));
  modal.show();
}

async function loadProfilePicture(userId, imgElement) {
  try {
    // If the image is already in cache, use it immediately
    if (profilePicCache.has(userId)) {
      imgElement.src = profilePicCache.get(userId);
      return;
    }

    const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      profilePicCache.set(userId, url);
      
      // Only update if the element still exists
      if (document.getElementById(imgElement.id)) {
        imgElement.src = url;
      }
    }
  } catch (error) {
    console.error(`Failed to load profile picture for user ${userId}:`, error);
  }
}

function loadContacts() {
  getTokenAndLoadMembers();
}
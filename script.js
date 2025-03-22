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

// Global variables
let accessToken = "";
let allMembers = [];
let activeDepartment = null;
let currentSort = "name";

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Initialize application
document.addEventListener('DOMContentLoaded', () => {
  console.log("DOM Content Loaded");
  
  // Initialize the sign-in modal
  const signInModal = new bootstrap.Modal(document.getElementById('signInModal'));
  
  // Handle redirect promise after login
  msalInstance.handleRedirectPromise().then(handleResponse).catch(err => {
    console.error("Redirect promise error:", err);
    // Show sign-in modal if there's an error
    signInModal.show();
  });

  // Set up event listeners
  setupEventListeners();
  
  // Check if we're returning from a redirect
  const currentAccounts = msalInstance.getAllAccounts();
  if (!currentAccounts || currentAccounts.length === 0) {
    // Show sign-in modal only if we're not returning from a redirect
    signInModal.show();
  }
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

function handleResponse(resp) {
  console.log("Handling response:", resp);
  if (resp !== null) {
    // If response is non-null, process it
    const account = resp.account;
    msalInstance.setActiveAccount(account);
    
    // Close the sign-in modal if it's open
    const signInModal = document.getElementById('signInModal');
    if (signInModal) {
      const modalInstance = bootstrap.Modal.getInstance(signInModal);
      if (modalInstance) {
        modalInstance.hide();
      } else {
        // If modal instance doesn't exist, create one and hide it
        const newModalInstance = new bootstrap.Modal(signInModal);
        newModalInstance.hide();
      }
      // Remove backdrop and modal-open class
      const backdrop = document.querySelector('.modal-backdrop');
      if (backdrop) backdrop.remove();
      document.body.classList.remove('modal-open');
      document.body.style.overflow = '';
      document.body.style.paddingRight = '';
    }
    
    showWelcomeUI(account);
    getTokenAndLoadMembers();
  } else {
    // If no response, check if we have any accounts
    const currentAccounts = msalInstance.getAllAccounts();
    if (!currentAccounts || currentAccounts.length === 0) {
      // No accounts, show sign in modal
      const signInModal = new bootstrap.Modal(document.getElementById('signInModal'));
      signInModal.show();
    } else {
      // Account exists, set active account and show welcome
      msalInstance.setActiveAccount(currentAccounts[0]);
      showWelcomeUI(currentAccounts[0]);
      getTokenAndLoadMembers();
    }
  }
}

function showWelcomeUI(account) {
  if (!account) return;
  
  const userGreeting = document.getElementById('userGreeting');
  userGreeting.textContent = `Hello, ${account.name}!`;
  userGreeting.classList.remove('d-none');
  
  // Show logout button
  document.getElementById('logoutBtn').classList.remove('d-none');
  
  // Close the modal if it's open
  const signInModal = document.getElementById('signInModal');
  const modalInstance = bootstrap.Modal.getInstance(signInModal);
  if (modalInstance) {
    modalInstance.hide();
    const backdrop = document.querySelector('.modal-backdrop');
    if (backdrop) backdrop.remove();
    document.body.classList.remove('modal-open');
  }
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
        throw new Error(`Failed to fetch members: ${response.status}`);
      }

      const data = await response.json();
      allMembersArray = allMembersArray.concat(data.value);
      
      document.getElementById("loading").textContent = `Loading members... (${allMembersArray.length} loaded)`;
      
      endpoint = data["@odata.nextLink"] || null;
    }

    allMembers = allMembersArray.sort((a, b) => a.displayName.localeCompare(b.displayName));
    renderDepartmentSelect(allMembers);
    renderContactList(allMembers);
    
    document.getElementById("loading").textContent = `Loaded ${allMembers.length} members`;
    document.getElementById("skeletonList").classList.add("d-none");
    setTimeout(() => {
      document.getElementById("loading").classList.add("d-none");
    }, 2000);

    loadProfilePictures(allMembers);

  } catch (error) {
    console.error("Error loading members:", error);
    document.getElementById("loading").textContent = "Failed to load all members.";
    document.getElementById("skeletonList").classList.add("d-none");
    alert("Failed to fetch group members.");
  }
}

async function loadProfilePictures(members) {
  for (const user of members) {
    try {
      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${user.id}/photo/$value`, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });

      if (response.ok) {
        const blob = await response.blob();
        const photoUrl = URL.createObjectURL(blob);
        user.photoUrl = photoUrl;

        // Update the profile picture in the UI
        const profilePics = document.querySelectorAll(`[data-user-id="${user.id}"]`);
        profilePics.forEach(pic => {
          pic.src = photoUrl;
          pic.alt = user.displayName;
        });
      }
    } catch (error) {
      console.log(`No photo available for ${user.displayName}`);
    }
  }
}

function renderContactList(members) {
  const list = document.getElementById("contactList");
  list.innerHTML = "";
  document.getElementById('totalContacts').textContent = members.length;
  document.getElementById('totalContactsLarge').textContent = members.length;

  if (members.length === 0) {
    list.innerHTML = '<p class="text-muted">No members found.</p>';
    return;
  }

  members.forEach(user => {
    const col = document.createElement("div");
    col.className = "col-12 col-md-6 col-lg-4";

    col.innerHTML = `
      <div class="card contact-card shadow-sm p-3 position-relative" data-id="${user.userPrincipalName}">
        <div class="quick-actions">
          <button class="btn btn-quick-action" onclick="event.stopPropagation(); copyToClipboard('${user.mail || user.userPrincipalName}')" title="Copy Email">
            <i class="bi bi-envelope"></i>
          </button>
          <button class="btn btn-quick-action" onclick="event.stopPropagation(); copyToClipboard('${user.mobilePhone || '-'}')" title="Copy Phone">
            <i class="bi bi-phone"></i>
          </button>
        </div>
        <div class="d-flex align-items-center mb-2">
          <img src="${user.photoUrl || 'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/icons/person-circle.svg'}"
               class="profile-pic me-3"
               data-user-id="${user.id}"
               alt="${user.displayName}">
          <div>
            <h5 class="card-title mb-1">${user.displayName}</h5>
            <span class="badge-department">${user.department || 'No Department'}</span>
          </div>
        </div>
        <div class="contact-details">
          <p class="card-text mb-1">
            <strong>Email:</strong> 
            <span class="copyable" onclick="event.stopPropagation(); copyToClipboard('${user.mail || user.userPrincipalName}')">${user.mail || user.userPrincipalName}</span>
          </p>
          <p class="card-text mb-1">
            <strong>Mobile:</strong> 
            <span class="copyable" onclick="event.stopPropagation(); copyToClipboard('${user.mobilePhone || '-'}')">${user.mobilePhone || '-'}</span>
          </p>
          <p class="card-text mb-0">
            <strong>Title:</strong> 
            <span class="copyable" onclick="event.stopPropagation(); copyToClipboard('${user.jobTitle || '-'}')">${user.jobTitle || '-'}</span>
          </p>
        </div>
      </div>
    `;

    col.querySelector(".contact-card").addEventListener("click", () => showDetails(user));
    list.appendChild(col);
  });
}

function showDetails(user) {
  const modalBody = document.getElementById("modalBodyContent");
  modalBody.innerHTML = `
    <div class="text-center mb-4">
      <img src="${user.photoUrl || 'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/icons/person-circle.svg'}"
           class="profile-pic-lg mb-3"
           data-user-id="${user.id}"
           alt="${user.displayName}">
      <h4>${user.displayName}</h4>
    </div>
    <p>
      <strong>Email:</strong> 
      <span class="copyable" onclick="copyToClipboard('${user.mail || user.userPrincipalName}')">${user.mail || user.userPrincipalName}</span>
    </p>
    <p>
      <strong>Mobile:</strong> 
      <span class="copyable" onclick="copyToClipboard('${user.mobilePhone || '-'}')">${user.mobilePhone || '-'}</span>
    </p>
    <p>
      <strong>Department:</strong> 
      <span class="copyable" onclick="copyToClipboard('${user.department || '-'}')">${user.department || '-'}</span>
    </p>
    <p>
      <strong>Job Title:</strong> 
      <span class="copyable" onclick="copyToClipboard('${user.jobTitle || '-'}')">${user.jobTitle || '-'}</span>
    </p>
    <p>
      <strong>Username:</strong> 
      <span class="copyable" onclick="copyToClipboard('${user.userPrincipalName}')">${user.userPrincipalName}</span>
    </p>
  `;

  const modal = new bootstrap.Modal(document.getElementById("detailsModal"));
  modal.show();
}

// Copy to clipboard function
function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => {
    const feedback = document.getElementById('copyFeedback');
    feedback.style.display = 'block';
    setTimeout(() => {
      feedback.style.display = 'none';
    }, 2000);
  });
}

// Update the signIn function
async function signIn() {
  try {
    console.log("Starting sign in process...");
    const loginRequest = {
      scopes: ["User.Read", "GroupMember.Read.All"],
      prompt: "select_account"
    };

    // Use redirect instead of popup for better compatibility
    console.log("Attempting redirect login...");
    await msalInstance.loginRedirect(loginRequest);
  } catch (error) {
    console.error("Error during sign in:", error);
    alert("Sign in failed. Please try again.");
  }
} 

/* global Office, Queue */

// IIFE wrapper to execute immediately
(function () {
  'use strict';

  console.log('=== TASKPANE.JS LOADING (IIFE START) ===');
  console.log('Timestamp:', new Date().toISOString());

  let eventCounter = 0;
  let isMonitoring = true;
  let monitoringInterval = null;
  let previousItemState = null;
  let isInitialized = false;
  let lastReadStatus = null;
  let isPinned = false;
  let contextSwitchCount = 0;
  let currentMailbox = null;
  let currentItemFrom = null;

  // Initialize Office
  Office.onReady((info) => {
    console.log('=== TASKPANE OFFICE.ONREADY FIRED ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);

    // Runtime checks
    console.log('=== RUNTIME CHECK ===');
    console.log('Supports Shared Runtime:', Office.context.requirements.isSetSupported('SharedRuntime', '1.1'));
    console.log('Mailbox version:', Office.context.mailbox.diagnostics.hostVersion);
    console.log('Host Name:', Office.context.mailbox.diagnostics.hostName);

    if (info.host === Office.HostType.Outlook) {
      // Use setTimeout to ensure DOM is ready
      setTimeout(() => {
        try {
          initializeTaskpane();
        } catch (error) {
          console.error('=== INITIALIZATION ERROR ===');
          console.error('Error:', error);
          console.error('Stack:', error.stack);
        }
      }, 100);
    }
  });

  function initializeTaskpane() {
    console.log('=== TASKPANE INITIALIZATION STARTED ===');
    console.log('Timestamp:', new Date().toISOString());

    try {
      // Verify DOM elements exist
      const clearLogBtn = document.getElementById('clear-log');
      const toggleMonitoringBtn = document.getElementById('toggle-monitoring');
      const triggerTestBtn = document.getElementById('trigger-test-event');
      const showAllMailboxesBtn = document.getElementById('show-all-mailboxes');

      console.log('DOM Elements Check:');
      console.log('  - clear-log:', clearLogBtn ? 'Found' : 'NOT FOUND');
      console.log('  - toggle-monitoring:', toggleMonitoringBtn ? 'Found' : 'NOT FOUND');
      console.log('  - trigger-test-event:', triggerTestBtn ? 'Found' : 'NOT FOUND');
      console.log('  - show-all-mailboxes:', showAllMailboxesBtn ? 'Found' : 'NOT FOUND');

      if (!clearLogBtn || !toggleMonitoringBtn || !triggerTestBtn) {
        throw new Error('Required DOM elements not found');
      }

      // Attach event handlers
      clearLogBtn.onclick = clearActivityLog;
      toggleMonitoringBtn.onclick = toggleMonitoring;
      triggerTestBtn.onclick = triggerTestEvent;

      if (showAllMailboxesBtn) {
        showAllMailboxesBtn.onclick = showMailboxActivity;
      }

      console.log('Event handlers attached successfully');

      // Check pinning status and provide guidance
      checkPinningStatus();

      logActivity('info', 'InboxAgent taskpane initialized');

      // Check for event runtime
      checkEventRuntime();

      // Load current user information
      loadCurrentUserInfo();

      // Update current item (this will also load the FROM address)
      updateCurrentItem();

      // Add Office event listeners
      addOfficeEventListeners();

      // Start deep monitoring immediately
      setTimeout(() => {
        startDeepMonitoring();
        logActivity('success', 'Deep monitoring started automatically');
      }, 500);

      // Set up persistence check
      setUpPersistenceMonitoring();

      isInitialized = true;

      console.log('=== INBOXAGENT TASKPANE INITIALIZED SUCCESSFULLY ===');
      console.log('Timestamp:', new Date().toISOString());
      console.log('Office Host:', Office.context.mailbox.diagnostics.hostName);
      console.log('Office Version:', Office.context.mailbox.diagnostics.hostVersion);
      console.log('Deep Monitoring: ACTIVE');

    } catch (error) {
      console.error('=== INITIALIZATION ERROR ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
      logActivity('error', `Initialization failed: ${error.message}`);
    }
  }

  function getCurrentMailboxInfo() {
    try {
      const userProfile = Office.context.mailbox.userProfile;

      let mailboxEmail = userProfile.emailAddress;
      let mailboxName = userProfile.displayName;

      return {
        email: mailboxEmail,
        name: mailboxName,
        displayText: `${mailboxName} <${mailboxEmail}>`
      };
    } catch (error) {
      console.error('Error getting mailbox info:', error);
      return {
        email: 'Unknown',
        name: 'Unknown',
        displayText: 'Unknown Mailbox'
      };
    }
  }

  function getCurrentItemFromInfo(callback) {
    const item = Office.context.mailbox.item;

    if (!item) {
      callback(null);
      return;
    }

    // For compose mode, there's no FROM (user is composing)
    if (item.itemType === Office.MailboxEnums.ItemType.Message &&
      item.itemClass && item.itemClass.includes('IPM.Note')) {

      getPropertyValue(item, 'from', (fromValue) => {
        if (fromValue) {
          callback({
            email: fromValue.emailAddress || 'Unknown',
            name: fromValue.displayName || fromValue.emailAddress || 'Unknown',
            displayText: fromValue.displayName ?
              `${fromValue.displayName} <${fromValue.emailAddress}>` :
              fromValue.emailAddress
          });
        } else {
          callback(null);
        }
      });
    } else {
      // Compose mode - user is the sender
      const userProfile = Office.context.mailbox.userProfile;
      callback({
        email: userProfile.emailAddress,
        name: userProfile.displayName,
        displayText: `${userProfile.displayName} <${userProfile.emailAddress}>`,
        isComposing: true
      });
    }
  }

  function updateItemFromDisplay() {
    getCurrentItemFromInfo((fromInfo) => {
      const itemFromElement = document.getElementById('item-from');

      if (itemFromElement) {
        if (fromInfo) {
          if (fromInfo.isComposing) {
            itemFromElement.textContent = `${fromInfo.name} (composing)`;
            itemFromElement.title = fromInfo.displayText;
          } else {
            itemFromElement.textContent = fromInfo.name;
            itemFromElement.title = fromInfo.displayText;
          }
          currentItemFrom = fromInfo;
          console.log('Item FROM updated:', fromInfo.displayText);
        } else {
          itemFromElement.textContent = 'N/A';
          itemFromElement.title = 'No sender information available';
          currentItemFrom = null;
        }
      }
    });
  }

  function checkPinningStatus() {
    console.log('=== CHECKING PINNING STATUS ===');
    console.log('Timestamp:', new Date().toISOString());

    try {
      // Check if already pinned (from localStorage)
      const userHasPinned = localStorage.getItem('inboxagent-pinned');

      if (userHasPinned === 'true') {
        // Hide the reminder
        const reminder = document.getElementById('pin-reminder');
        if (reminder) {
          reminder.classList.add('hidden');
        }
        console.log('User has previously pinned - reminder hidden');
        logActivity('success', '✓ Taskpane is pinned!');
      } else {
        // Show the reminder and set up the "Got it" button
        const gotItBtn = document.getElementById('got-it-pin');
        if (gotItBtn) {
          gotItBtn.onclick = () => {
            // User acknowledged - hide reminder
            const reminder = document.getElementById('pin-reminder');
            if (reminder) {
              reminder.style.transition = 'all 0.3s ease-out';
              reminder.style.opacity = '0';
              reminder.style.transform = 'scale(0.95)';

              setTimeout(() => {
                reminder.classList.add('hidden');
              }, 300);
            }

            // Save preference
            localStorage.setItem('inboxagent-pinned', 'true');

            logActivity('success', '👍 Great! Now click the 📌 pin icon in the top-right corner');

            console.log('⭐ USER ACKNOWLEDGED PIN REQUEST');
          };
        }
      }

      // Check if UI context is available
      if (Office.context.ui) {
        console.log('UI Context available');

        // Try to detect if pinned (read-only)
        if (typeof Office.context.ui.isPinned !== 'undefined') {
          isPinned = Office.context.ui.isPinned;
          console.log('Pinning status:', isPinned ? 'PINNED' : 'NOT PINNED');

          if (isPinned) {
            // User has pinned! Save this
            localStorage.setItem('inboxagent-pinned', 'true');

            // Hide reminder
            const reminder = document.getElementById('pin-reminder');
            if (reminder) {
              reminder.classList.add('hidden');
            }

            logActivity('success', '✓ Taskpane is pinned!');
          }
        } else {
          console.log('Pinning status not available via API');
        }
      }

      // If not acknowledged yet, show helpful messages
      if (userHasPinned !== 'true') {
        setTimeout(() => {
          logActivity('info', '💡 Click the 📌 pin icon to keep this pane open');
          logActivity('info', '📌 Pinning keeps the taskpane visible when composing emails');
        }, 2000);
      }

    } catch (error) {
      console.error('Error checking pinning status:', error);
    }
  }

  function setUpPersistenceMonitoring() {
    console.log('=== SETTING UP PERSISTENCE MONITORING ===');

    // Monitor if taskpane stays visible across context switches
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => {
          contextSwitchCount++;
          console.log('Context switch detected. Count:', contextSwitchCount);

          if (!document.hidden && contextSwitchCount > 1) {
            // Taskpane stayed visible through context switch = likely pinned!
            console.log('✓ TASKPANE PERSISTED THROUGH CONTEXT SWITCH');
            console.log('Context switches:', contextSwitchCount);

            if (!isPinned) {
              isPinned = true;
              localStorage.setItem('inboxagent-pinned', 'true');

              // Hide reminder
              const reminder = document.getElementById('pin-reminder');
              if (reminder && !reminder.classList.contains('hidden')) {
                reminder.style.transition = 'all 0.3s ease-out';
                reminder.style.opacity = '0';
                reminder.style.transform = 'scale(0.95)';

                setTimeout(() => {
                  reminder.classList.add('hidden');
                }, 300);

                logActivity('success', '🎉 Taskpane is now pinned and will stay visible!');
              }
            }
          } else if (document.hidden && contextSwitchCount > 0) {
            console.log('⚠ Taskpane hidden after context switch - not pinned');
            logActivity('warning', 'Taskpane was closed. Pin it to keep it visible!');
          }
        },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ ItemChanged handler for persistence monitoring attached');
          } else {
            console.error('Failed to attach ItemChanged handler:', result.error);
          }
        }
      );
    }

    // Monitor visibility changes
    let visibilityCheckCount = 0;
    const visibilityInterval = setInterval(() => {
      visibilityCheckCount++;

      if (!document.hidden) {
        // Taskpane is visible
        if (visibilityCheckCount % 20 === 0) { // Log every 20 checks (20 seconds)
          console.log('✓ Taskpane visibility check:', visibilityCheckCount, '- VISIBLE');
        }
      } else {
        console.log('⚠ Taskpane is hidden');
      }

      // Stop checking after 5 minutes
      if (visibilityCheckCount > 300) {
        clearInterval(visibilityInterval);
        console.log('Stopped visibility monitoring after 5 minutes');
      }
    }, 1000);

    // Listen for visibility change events
    document.addEventListener('visibilitychange', () => {
      if (document.hidden) {
        console.log('=== TASKPANE HIDDEN ===');
        console.log('Timestamp:', new Date().toISOString());

        const userHasPinned = localStorage.getItem('inboxagent-pinned');
        if (userHasPinned !== 'true') {
          logActivity('warning', 'Taskpane hidden - pin it to keep visible');
        }
      } else {
        console.log('=== TASKPANE VISIBLE ===');
        console.log('Timestamp:', new Date().toISOString());
        logActivity('success', 'Taskpane is now visible');

        // Refresh monitoring when taskpane becomes visible again
        if (isMonitoring && !monitoringInterval) {
          startDeepMonitoring();
        }
      }
    });

    // Listen for beforeunload (taskpane closing)
    window.addEventListener('beforeunload', () => {
      console.log('=== TASKPANE CLOSING ===');
      console.log('Timestamp:', new Date().toISOString());
      console.log('Events tracked:', eventCounter);
      console.log('Context switches:', contextSwitchCount);

      // Clean up
      if (monitoringInterval) {
        clearInterval(monitoringInterval);
      }
    });

    // Listen for page show/hide (back/forward navigation)
    window.addEventListener('pageshow', (event) => {
      if (event.persisted) {
        console.log('=== TASKPANE RESTORED FROM CACHE ===');
        logActivity('info', 'Taskpane restored');

        // Reinitialize monitoring
        if (isMonitoring) {
          startDeepMonitoring();
        }
      }
    });

    console.log('✓ Persistence monitoring configured');
  }

  function loadCurrentUserInfo() {
    console.log('=== LOADING USER INFORMATION ===');
    console.log('Timestamp:', new Date().toISOString());

    try {
      const userProfile = Office.context.mailbox.userProfile;
      const diagnostics = Office.context.mailbox.diagnostics;
      const mailbox = Office.context.mailbox;

      if (userProfile) {
        console.log('=== COMPLETE USER PROFILE ===');

        // Basic Info
        console.log('Display Name:', userProfile.displayName);
        console.log('Email Address:', userProfile.emailAddress);
        console.log('Account Type:', userProfile.accountType);
        console.log('Time Zone:', userProfile.timeZone);

        // Store current mailbox info
        currentMailbox = {
          email: userProfile.emailAddress,
          name: userProfile.displayName,
          accountType: userProfile.accountType
        };

        // Update UI
        document.getElementById('user-display-name').textContent = userProfile.displayName || 'N/A';
        document.getElementById('user-email').textContent = userProfile.emailAddress || 'N/A';
        document.getElementById('user-account-type').textContent = userProfile.accountType || 'N/A';
        document.getElementById('user-timezone').textContent = userProfile.timeZone || 'N/A';

        console.log('=== MAILBOX DIAGNOSTICS ===');
        console.log('Host Name:', diagnostics.hostName);
        console.log('Host Version:', diagnostics.hostVersion);

        // Convert REST endpoint to readable format
        const restUrl = mailbox.restUrl;
        console.log('REST URL:', restUrl);

        // EWS endpoint
        const ewsUrl = mailbox.ewsUrl;
        console.log('EWS URL:', ewsUrl);

        // Additional diagnostic info
        if (diagnostics.OWAView) {
          console.log('OWA View:', diagnostics.OWAView);
        }

        console.log('=== MAILBOX SETTINGS ===');

        // Get mailbox settings (require REST call for full details)
        const settings = Office.context.roamingSettings;
        console.log('Roaming Settings Available:', settings !== null);

        // Log activity with full user context
        const userContext = `${userProfile.displayName} (${userProfile.emailAddress})`;
        logActivity('success', `User loaded: ${userContext}`);
        logActivity('info', `Account Type: ${userProfile.accountType}`);
        logActivity('info', `Time Zone: ${userProfile.timeZone}`);

        console.log('=== USER INFORMATION LOADED SUCCESSFULLY ===');

        return {
          displayName: userProfile.displayName,
          emailAddress: userProfile.emailAddress,
          accountType: userProfile.accountType,
          timeZone: userProfile.timeZone,
          hostName: diagnostics.hostName,
          hostVersion: diagnostics.hostVersion,
          restUrl: restUrl,
          ewsUrl: ewsUrl
        };

      } else {
        throw new Error('User profile not available');
      }

    } catch (error) {
      console.error('=== ERROR LOADING USER INFORMATION ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
      logActivity('error', `Failed to load user info: ${error.message}`);

      // Set error states
      ['user-display-name', 'user-email', 'user-account-type', 'user-timezone'].forEach(id => {
        const element = document.getElementById(id);
        if (element) {
          element.textContent = 'Error';
          element.classList.add('inactive');
        }
      });

      return null;
    }
  }

  function checkEventRuntime() {
    console.log('=== CHECKING EVENT RUNTIME ===');
    console.log('Timestamp:', new Date().toISOString());

    const runtimeStatus = document.getElementById('runtime-status');

    if (!runtimeStatus) {
      console.error('runtime-status element not found');
      return;
    }

    if (Office.context.mailbox.item && Office.context.mailbox.addHandlerAsync) {
      runtimeStatus.textContent = 'Active';
      runtimeStatus.classList.add('active');
      logActivity('success', 'Event-based activation runtime is active');
      console.log('Event-based activation is supported');
    } else {
      runtimeStatus.textContent = 'Not Available';
      runtimeStatus.classList.add('inactive');
      logActivity('warning', 'Event-based activation not available');
      console.log('Event-based activation is NOT supported');
    }
  }

  // Helper function to get property value (handles both read and compose modes)
  function getPropertyValue(item, propertyName, callback) {
    if (!item) {
      console.log(`getPropertyValue: No item provided for ${propertyName}`);
      callback(null);
      return;
    }

    const property = item[propertyName];

    if (!property) {
      console.log(`getPropertyValue: Property ${propertyName} not found on item`);
      callback(null);
      return;
    }

    // Check if it's a compose mode property (has getAsync)
    if (typeof property.getAsync === 'function') {
      console.log(`getPropertyValue: Using getAsync for ${propertyName}`);
      try {
        property.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`getPropertyValue: Got value for ${propertyName}:`, result.value);
            callback(result.value);
          } else {
            console.error(`getPropertyValue: Failed to get ${propertyName}:`, result.error);
            callback(null);
          }
        });
      } catch (error) {
        console.error(`getPropertyValue: Exception getting ${propertyName}:`, error);
        callback(null);
      }
    } else {
      // Read mode - direct property access
      console.log(`getPropertyValue: Direct access for ${propertyName}:`, property);
      callback(property);
    }
  }

  function triggerTestEvent() {
    console.log('=== TRIGGERING TEST EVENT ===');
    console.log('Timestamp:', new Date().toISOString());

    const mailboxInfo = getCurrentMailboxInfo();
    console.log('Current User Mailbox:', mailboxInfo);

    logActivity('info', 'Test event triggered - check console for details');

    const item = Office.context.mailbox.item;
    if (item) {
      console.log('Current Item Details:');
      console.log('  - Item Type:', item.itemType);
      console.log('  - Item Class:', item.itemClass);
      console.log('  - Item ID:', item.itemId || 'No ID (new item)');
      console.log('  - Conversation ID:', item.conversationId);
      console.log('  - Read Status:', item.read);
      console.log('  - User Mailbox:', mailboxInfo.displayText);

      const testQueue = new Queue({ results: [], concurrency: 1 });

      testQueue.push(cb => {
        getPropertyValue(item, 'subject', (value) => {
          console.log('  - Subject:', value);
          logActivity('info', `Subject: ${value}`);
          cb();
        });
      });

      testQueue.push(cb => {
        getPropertyValue(item, 'from', (value) => {
          console.log('  - From (Sender):', JSON.stringify(value, null, 2));
          if (value) {
            logActivity('info', `From: ${value.displayName || value.emailAddress}`);
          }
          cb();
        });
      });

      testQueue.push(cb => {
        getPropertyValue(item, 'to', (value) => {
          console.log('  - To:', JSON.stringify(value, null, 2));
          cb();
        });
      });

      testQueue.push(cb => {
        getPropertyValue(item, 'categories', (value) => {
          console.log('  - Categories:', JSON.stringify(value, null, 2));
          logActivity('info', `Categories: ${JSON.stringify(value)}`);
          cb();
        });
      });

      if (item.attachments) {
        testQueue.push(cb => {
          console.log('  - Attachments:', item.attachments.length);
          item.attachments.forEach(att => {
            console.log(`    * ${att.name} (${att.size} bytes)`);
          });
          cb();
        });
      }

      testQueue.start((err) => {
        if (err) {
          console.error('Test queue error:', err);
        } else {
          console.log('Test queue completed successfully');
        }
      });
    } else {
      console.log('No item currently selected');
      logActivity('warning', 'No item currently selected');
    }
  }

  function addOfficeEventListeners() {
    console.log('=== ADDING OFFICE EVENT LISTENERS IN TASKPANE ===');
    console.log('Timestamp:', new Date().toISOString());

    try {
      // Item Changed Event
      if (Office.context.mailbox.addHandlerAsync) {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          onItemChanged,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              logActivity('success', 'ItemChanged listener attached');
              console.log('=== EVENT LISTENER ATTACHED ===');
              console.log('Event Type: ItemChanged');
              console.log('Timestamp:', new Date().toISOString());
            } else {
              logActivity('error', 'Failed to attach ItemChanged listener');
              console.error('=== EVENT LISTENER FAILED ===');
              console.error('Event Type: ItemChanged');
              console.error('Error:', result.error);
            }
          }
        );
      }

      // Recipients Changed Event (if in compose mode)
      const item = Office.context.mailbox.item;
      if (item && item.addHandlerAsync) {
        const eventTypes = [
          'RecipientsChanged',
          'RecurrenceChanged',
          'AppointmentTimeChanged'
        ];

        eventTypes.forEach(eventType => {
          if (Office.EventType[eventType]) {
            item.addHandlerAsync(
              Office.EventType[eventType],
              (args) => onItemPropertyChanged(eventType, args),
              (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  logActivity('success', `${eventType} listener attached`);
                  console.log(`=== EVENT LISTENER ATTACHED: ${eventType} ===`);
                }
              }
            );
          }
        });
      }

      console.log('=== FINISHED ADDING OFFICE EVENT LISTENERS ===');
    } catch (error) {
      console.error('=== ERROR ADDING EVENT LISTENERS ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
    }
  }

  function onItemChanged(args) {
    console.log('=== ITEM CHANGED EVENT FIRED (TASKPANE) ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Args:', JSON.stringify(args, null, 2));

    logActivity('info', `Item changed - Loading new item details`);

    eventCounter++;
    updateEventCounter();
    updateCurrentItem();

    // Reset monitoring state for new item
    previousItemState = null;
    lastReadStatus = null;
    if (isMonitoring) {
      captureCurrentItemState();
    }
  }

  function onItemPropertyChanged(eventType, args) {
    console.log(`=== ${eventType.toUpperCase()} EVENT FIRED (TASKPANE) ===`);
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Args:', JSON.stringify(args, null, 2));

    logActivity('warning', `${eventType} detected`);
    eventCounter++;
    updateEventCounter();
  }

  function updateCurrentItem() {
    console.log('=== UPDATING CURRENT ITEM ===');

    const currentItemElement = document.getElementById('current-item');
    if (!currentItemElement) {
      console.error('current-item element not found');
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item) {
      console.log('No item available');
      currentItemElement.textContent = 'No item selected';

      // Clear the FROM display
      const itemFromElement = document.getElementById('item-from');
      if (itemFromElement) {
        itemFromElement.textContent = 'N/A';
      }
      return;
    }

    console.log('Item available, getting subject and from...');

    getPropertyValue(item, 'subject', (subject) => {
      const displaySubject = subject || '(No Subject)';
      currentItemElement.textContent =
        displaySubject.substring(0, 30) + (displaySubject.length > 30 ? '...' : '');

      const mailboxInfo = getCurrentMailboxInfo();
      console.log('=== CURRENT ITEM UPDATED ===');
      console.log('Subject:', displaySubject);
      console.log('Item Type:', item.itemType);
      console.log('Item ID:', item.itemId || 'New item (no ID)');
      console.log('User Mailbox:', mailboxInfo.displayText);

      // Update the FROM display
      updateItemFromDisplay();
    });
  }

  function toggleMonitoring() {
    console.log('=== TOGGLE MONITORING CLICKED ===');

    isMonitoring = !isMonitoring;
    const button = document.getElementById('toggle-monitoring');
    const statusElement = document.getElementById('monitoring-status');

    if (!button || !statusElement) {
      console.error('Button or status element not found');
      return;
    }

    if (isMonitoring) {
      button.textContent = 'Pause Monitoring';
      button.classList.remove('btn-success');
      button.classList.add('btn-warning');
      statusElement.textContent = 'Active';
      statusElement.classList.remove('paused');
      statusElement.classList.add('active');
      startDeepMonitoring();
      logActivity('success', 'Deep monitoring resumed');
    } else {
      button.textContent = 'Resume Monitoring';
      button.classList.remove('btn-warning');
      button.classList.add('btn-success');
      statusElement.textContent = 'Paused';
      statusElement.classList.remove('active');
      statusElement.classList.add('paused');
      stopDeepMonitoring();
      logActivity('warning', 'Deep monitoring paused');
    }

    console.log('=== MONITORING TOGGLED XXX ===');
    console.log('Monitoring Active:', isMonitoring);
    console.log('Timestamp:', new Date().toISOString());
  }

  function startDeepMonitoring() {
    console.log('=== STARTING DEEP MONITORING XXX ===');

    try {
      captureCurrentItemState();

      // Initialize read status
      const item = Office.context.mailbox.item;
      if (item && typeof item.read !== 'undefined') {
        lastReadStatus = item.read;
      }

      // Poll for changes every 1 second
      if (monitoringInterval) {
        clearInterval(monitoringInterval);
      }

      monitoringInterval = setInterval(() => {
        checkForItemChanges();
        monitorReadStatusChanges();
      }, 1000);

      console.log('=== DEEP MONITORING STARTED ===');
      console.log('Polling Interval: 1000ms');
      console.log('Timestamp:', new Date().toISOString());
    } catch (error) {
      console.error('=== ERROR STARTING MONITORING ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
    }
  }

  function stopDeepMonitoring() {
    if (monitoringInterval) {
      clearInterval(monitoringInterval);
      monitoringInterval = null;
    }
    previousItemState = null;
    lastReadStatus = null;

    console.log('=== DEEP MONITORING STOPPED XXX ===');
    console.log('Timestamp:', new Date().toISOString());
  }

  // Monitor read status changes separately
  function monitorReadStatusChanges() {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.read === 'undefined') return;

    const currentReadStatus = item.read;

    if (lastReadStatus !== null && lastReadStatus !== currentReadStatus) {
      console.log('=== READ STATUS CHANGED ===');
      console.log('Previous:', lastReadStatus ? 'Read' : 'Unread');
      console.log('Current:', currentReadStatus ? 'Read' : 'Unread');
      console.log('Email:', item.subject || 'Unknown');

      const statusChange = currentReadStatus ? 'marked as Read' : 'marked as Unread';

      // Include FROM info if available
      if (currentItemFrom) {
        logActivity('warning', `Email from ${currentItemFrom.name} ${statusChange}`);
      } else {
        logActivity('warning', `Email ${statusChange}`);
      }

      eventCounter++;
      updateEventCounter();
    }

    lastReadStatus = currentReadStatus;
  }

  function captureCurrentItemState() {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.log('captureCurrentItemState: No item available');
      return;
    }

    console.log('=== CAPTURING ITEM STATE ===');

    const mailboxInfo = getCurrentMailboxInfo();
    const captureQueue = new Queue({ results: [], concurrency: 1 });
    const state = {
      itemType: item.itemType,
      itemId: item.itemId,
      itemClass: item.itemClass || null,
      userMailbox: mailboxInfo.email,
      userMailboxName: mailboxInfo.name
    };

    // Capture subject
    captureQueue.push(cb => {
      getPropertyValue(item, 'subject', (value) => {
        state.subject = value;
        cb();
      });
    });

    // Capture read status
    captureQueue.push(cb => {
      if (typeof item.read !== 'undefined') {
        state.read = item.read;
        console.log('Read status captured:', item.read);
      } else {
        state.read = null;
      }
      cb();
    });

    // Capture categories
    captureQueue.push(cb => {
      getPropertyValue(item, 'categories', (value) => {
        state.categories = value;
        cb();
      });
    });

    // Capture internet message id
    if (item.internetMessageId) {
      captureQueue.push(cb => {
        state.internetMessageId = item.internetMessageId;
        cb();
      });
    }

    // Capture conversation id
    if (item.conversationId) {
      captureQueue.push(cb => {
        state.conversationId = item.conversationId;
        cb();
      });
    }

    // Capture from (SENDER of the email)
    captureQueue.push(cb => {
      getPropertyValue(item, 'from', (value) => {
        if (value) {
          state.from = value;
          state.fromEmail = value.emailAddress;
          state.fromName = value.displayName || value.emailAddress;
        }
        cb();
      });
    });

    // Capture to recipients
    captureQueue.push(cb => {
      getPropertyValue(item, 'to', (value) => {
        state.to = value;
        cb();
      });
    });

    // Capture cc recipients
    captureQueue.push(cb => {
      getPropertyValue(item, 'cc', (value) => {
        state.cc = value;
        cb();
      });
    });

    // Capture attachments
    if (item.attachments) {
      captureQueue.push(cb => {
        state.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType
        }));
        cb();
      });
    }

    captureQueue.push(cb => {
      previousItemState = state;
      console.log('=== ITEM STATE CAPTURED ===');
      console.log('Captured at:', new Date().toISOString());
      console.log('User Mailbox:', mailboxInfo.displayText);
      console.log('Email From:', state.fromName || 'N/A');
      console.log('State:', JSON.stringify(state, null, 2));
      cb();
    });

    captureQueue.start((err) => {
      if (err) {
        console.error('Capture queue error:', err);
      } else {
        console.log('Capture queue completed successfully');
      }
    });
  }

  function checkForItemChanges() {
    if (!previousItemState) {
      captureCurrentItemState();
      return;
    }

    const item = Office.context.mailbox.item;

    // Check if item disappeared (marked as junk and moved)
    if (!item && previousItemState.itemId) {
      console.log('=== ITEM DISAPPEARED ===');
      console.log('Previous item:', previousItemState.subject);
      console.log('From:', previousItemState.fromName);
      console.log('User mailbox:', previousItemState.userMailbox);
      logActivity('error', `Email from ${previousItemState.fromName || 'Unknown'} disappeared - possibly marked as junk or deleted`);

      previousItemState = null;
      return;
    }

    if (!item) return;

    const mailboxInfo = getCurrentMailboxInfo();
    const checkQueue = new Queue({ results: [], concurrency: 1 });
    const currentState = {
      itemType: item.itemType,
      itemId: item.itemId,
      itemClass: item.itemClass || null,
      userMailbox: mailboxInfo.email,
      userMailboxName: mailboxInfo.name
    };

    checkQueue.push(cb => {
      getPropertyValue(item, 'subject', (value) => {
        currentState.subject = value;
        cb();
      });
    });

    // Capture current read status
    checkQueue.push(cb => {
      if (typeof item.read !== 'undefined') {
        currentState.read = item.read;
      } else {
        currentState.read = null;
      }
      cb();
    });

    checkQueue.push(cb => {
      getPropertyValue(item, 'categories', (value) => {
        currentState.categories = value;
        cb();
      });
    });

    if (item.internetMessageId) {
      checkQueue.push(cb => {
        currentState.internetMessageId = item.internetMessageId;
        cb();
      });
    }

    if (item.conversationId) {
      checkQueue.push(cb => {
        currentState.conversationId = item.conversationId;
        cb();
      });
    }

    checkQueue.push(cb => {
      getPropertyValue(item, 'from', (value) => {
        if (value) {
          currentState.from = value;
          currentState.fromEmail = value.emailAddress;
          currentState.fromName = value.displayName || value.emailAddress;
        }
        cb();
      });
    });

    checkQueue.push(cb => {
      getPropertyValue(item, 'to', (value) => {
        currentState.to = value;
        cb();
      });
    });

    checkQueue.push(cb => {
      getPropertyValue(item, 'cc', (value) => {
        currentState.cc = value;
        cb();
      });
    });

    if (item.attachments) {
      checkQueue.push(cb => {
        currentState.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType
        }));
        cb();
      });
    }

    // Compare states
    checkQueue.push(cb => {
      compareStates(previousItemState, currentState);

      // Check for junk marking
      detectJunkMarking(previousItemState, currentState);

      // Check for folder changes (simple approach)
      detectFolderChangeSimple(previousItemState, currentState);

      previousItemState = currentState;
      cb();
    });

    checkQueue.start((err) => {
      if (err) {
        console.error('Check queue error:', err);
      }
    });
  }

  function compareStates(oldState, newState) {
    const oldJSON = JSON.stringify(oldState);
    const newJSON = JSON.stringify(newState);

    if (oldJSON !== newJSON) {
      console.log('=== ITEM STATE CHANGED ===');
      console.log('Comparison time:', new Date().toISOString());
      console.log('User Mailbox:', newState.userMailbox);
      console.log('Email From:', newState.fromName || 'N/A');
      console.log('Previous State:', oldJSON);
      console.log('Current State:', newJSON);

      const changes = [];

      // Check subject change
      if (oldState.subject !== newState.subject) {
        const change = `Subject: "${oldState.subject}" → "${newState.subject}"`;
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check read status change
      if (oldState.read !== newState.read) {
        const oldStatus = oldState.read ? 'Read' : 'Unread';
        const newStatus = newState.read ? 'Read' : 'Unread';
        const change = `Read Status: ${oldStatus} → ${newStatus}`;
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check categories change
      const oldCategories = JSON.stringify(oldState.categories || []);
      const newCategories = JSON.stringify(newState.categories || []);
      if (oldCategories !== newCategories) {
        const change = `Categories: ${oldCategories} → ${newCategories}`;
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check To recipients change
      const oldTo = JSON.stringify(oldState.to || []);
      const newTo = JSON.stringify(newState.to || []);
      if (oldTo !== newTo) {
        const change = 'To recipients changed';
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check CC recipients change
      const oldCc = JSON.stringify(oldState.cc || []);
      const newCc = JSON.stringify(newState.cc || []);
      if (oldCc !== newCc) {
        const change = 'CC recipients changed';
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check From change (sender)
      const oldFrom = JSON.stringify(oldState.from || null);
      const newFrom = JSON.stringify(newState.from || null);
      if (oldFrom !== newFrom) {
        const change = `From: ${oldState.fromName || 'Unknown'} → ${newState.fromName || 'Unknown'}`;
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check attachments change
      const oldAttachments = JSON.stringify(oldState.attachments || []);
      const newAttachments = JSON.stringify(newState.attachments || []);
      if (oldAttachments !== newAttachments) {
        const oldCount = oldState.attachments ? oldState.attachments.length : 0;
        const newCount = newState.attachments ? newState.attachments.length : 0;
        const change = `Attachments: ${oldCount} → ${newCount}`;
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);
      }

      // Check item ID change (different email selected)
      if (oldState.itemId !== newState.itemId) {
        const change = 'Different item selected';
        changes.push(change);
        logActivity('info', change);
        console.log('✓', change);
      }

      // Check item class change (can indicate junk marking)
      if (oldState.itemClass !== newState.itemClass) {
        const change = `Item Class: ${oldState.itemClass} → ${newState.itemClass}`;
        changes.push(change);
        logActivity('warning', change);
        console.log('✓', change);

        if (newState.itemClass?.includes('SMIME') || newState.itemClass?.includes('Rules')) {
          logActivity('error', 'Email may have been marked as junk or processed by rules');
          console.log('=== POSSIBLE JUNK MARKING DETECTED ===');
        }
      }

      // Check conversation change (possible reply/forward)
      if (oldState.conversationId !== newState.conversationId) {
        changes.push('Conversation changed');
        detectEmailAction(oldState, newState);
      } else if (oldState.itemId !== newState.itemId &&
        oldState.conversationId === newState.conversationId) {
        // Same conversation but different item = reply or forward
        detectEmailAction(oldState, newState);
      }

      if (changes.length > 0) {
        console.log(`✓ Total changes detected: ${changes.length}`);
        eventCounter++;
        updateEventCounter();
      } else {
        console.log('⚠ JSON differs but no specific property changes found');
        console.log('This might be due to object property ordering or other differences');
      }
    } else {
      // Only log occasionally to reduce console spam
      if (Math.random() < 0.02) { // 2% of the time
        console.log('✓ No state changes detected (polling...)');
      }
    }
  }

  function detectEmailAction(oldState, newState) {
    // Detect reply or forward actions
    if (oldState.conversationId && newState.conversationId) {
      if (oldState.conversationId === newState.conversationId &&
        oldState.itemId !== newState.itemId) {

        console.log('=== EMAIL ACTION DETECTED ===');
        console.log('Action Type: REPLY or FORWARD');
        console.log('Original Item ID:', oldState.itemId);
        console.log('New Item ID:', newState.itemId);
        console.log('Conversation ID:', newState.conversationId);
        console.log('Original Subject:', oldState.subject);
        console.log('New Subject:', newState.subject);
        console.log('Original From:', oldState.fromName);
        console.log('User Mailbox:', newState.userMailbox);

        let actionType = 'UNKNOWN';
        if (newState.subject && oldState.subject) {
          if (newState.subject.startsWith('RE:') || newState.subject.startsWith('Re:')) {
            actionType = 'REPLY';
          } else if (newState.subject.startsWith('FW:') || newState.subject.startsWith('Fw:')) {
            actionType = 'FORWARD';
          }
        }

        logActivity('success', `${actionType} to email from ${oldState.fromName || 'Unknown'}: "${oldState.subject}"`);

        console.log('Detected Action:', actionType);
      }
    }
  }

  // Detect if email was marked as junk by monitoring item disappearance
  function detectJunkMarking(oldState, newState) {
    // Case 1: Item ID changed but we're still in the same context
    if (oldState.itemId && newState.itemId && oldState.itemId !== newState.itemId) {
      console.log('Item ID changed - email may have been moved');
      logActivity('warning', 'Email moved or marked as junk/not junk');
    }

    // Case 2: Item became null (disappeared)
    if (oldState.itemId && !newState.itemId) {
      console.log('Item disappeared - likely moved to Junk or Deleted');
      logActivity('error', `Email from ${oldState.fromName || 'Unknown'} disappeared - possibly marked as junk`);

      console.log('=== EMAIL MARKED AS JUNK (LIKELY) ===');
      console.log('Subject:', oldState.subject);
      console.log('From Name:', oldState.fromName);
      console.log('From Email:', oldState.fromEmail);
      console.log('Item ID:', oldState.itemId);
      console.log('User Mailbox:', oldState.userMailbox);
    }
  }

  // Simplified folder change detection (no EWS needed)
  function detectFolderChangeSimple(oldState, newState) {
    // If the item ID changed but conversation is the same, it might have moved
    if (oldState.itemId && newState.itemId && oldState.itemId !== newState.itemId) {
      // Check if it's the same conversation (not a reply/forward)
      if (oldState.conversationId === newState.conversationId) {
        // Check if subject is exactly the same (not "RE:" or "FW:")
        if (oldState.subject === newState.subject) {
          console.log('=== POSSIBLE FOLDER MOVE DETECTED ===');
          console.log('Item IDs differ but conversation and subject same');
          console.log('Old Item ID:', oldState.itemId);
          console.log('New Item ID:', newState.itemId);
          console.log('Subject:', newState.subject);
          console.log('From:', newState.fromName);
          console.log('User Mailbox:', newState.userMailbox);

          logActivity('warning', `Email from ${newState.fromName || 'Unknown'} may have been moved: "${newState.subject}"`);

          eventCounter++;
          updateEventCounter();

          return true;
        }
      }
    }

    return false;
  }

  function logActivity(type, message) {
    try {
      const activityLog = document.getElementById('activity-log');
      if (!activityLog) {
        console.error('activity-log element not found');
        return;
      }

      const mailboxInfo = getCurrentMailboxInfo();

      const activityItem = document.createElement('div');
      activityItem.className = `activity-item ${type}`;

      const time = document.createElement('div');
      time.className = 'activity-time';
      time.textContent = new Date().toLocaleTimeString();

      const msg = document.createElement('div');
      msg.className = 'activity-message';

      // Include user mailbox in message with visual indicator
      const mailboxShort = mailboxInfo.email.split('@')[0]; // Get part before @
      msg.innerHTML = `<strong style="color: #667eea;">[${mailboxShort}]</strong> ${message}`;

      activityItem.appendChild(time);
      activityItem.appendChild(msg);

      // Insert at the top
      if (activityLog.firstChild) {
        activityLog.insertBefore(activityItem, activityLog.firstChild);
      } else {
        activityLog.appendChild(activityItem);
      }

      // Keep only last 50 items
      while (activityLog.children.length > 50) {
        activityLog.removeChild(activityLog.lastChild);
      }

      // Store event to history with mailbox and from info
      const fromInfo = currentItemFrom ? {
        email: currentItemFrom.email,
        name: currentItemFrom.name
      } : null;

      storeEventToHistory(mailboxInfo.email, mailboxInfo.name, fromInfo, type, message);

    } catch (error) {
      console.error('Error logging activity:', error);
    }
  }

  function storeEventToHistory(userMailbox, userMailboxName, fromInfo, type, message) {
    try {
      let history = JSON.parse(localStorage.getItem('inboxagent-history') || '[]');

      history.push({
        timestamp: new Date().toISOString(),
        userMailbox: userMailbox,
        userMailboxName: userMailboxName,
        fromEmail: fromInfo ? fromInfo.email : null,
        fromName: fromInfo ? fromInfo.name : null,
        type: type,
        message: message
      });

      // Keep last 500 events
      if (history.length > 500) {
        history = history.slice(-500);
      }

      localStorage.setItem('inboxagent-history', JSON.stringify(history));
    } catch (error) {
      console.error('Error storing event history:', error);
    }
  }

  function showMailboxActivity() {
    try {
      const history = JSON.parse(localStorage.getItem('inboxagent-history') || '[]');

      if (history.length === 0) {
        logActivity('info', 'No activity history available yet');
        console.log('No activity history available');
        return;
      }

      // Group by user mailbox
      const byMailbox = history.reduce((acc, event) => {
        const key = event.userMailbox;
        if (!acc[key]) {
          acc[key] = {
            name: event.userMailboxName || event.userMailbox,
            events: []
          };
        }
        acc[key].events.push(event);
        return acc;
      }, {});

      // Also group by email sender
      const bySender = history.filter(e => e.fromEmail).reduce((acc, event) => {
        const key = event.fromEmail;
        if (!acc[key]) {
          acc[key] = {
            name: event.fromName || event.fromEmail,
            events: []
          };
        }
        acc[key].events.push(event);
        return acc;
      }, {});

      console.log('=== ACTIVITY BY USER MAILBOX ===');
      console.log('Total events:', history.length);
      console.log('User Mailboxes tracked:', Object.keys(byMailbox).length);
      console.log('Email Senders tracked:', Object.keys(bySender).length);
      console.log('');

      Object.keys(byMailbox).forEach(mailbox => {
        const data = byMailbox[mailbox];
        console.log(`📬 User Mailbox: ${data.name} (${mailbox})`);
        console.log(`   Total events: ${data.events.length}`);

        // Count by type
        const byType = data.events.reduce((acc, event) => {
          acc[event.type] = (acc[event.type] || 0) + 1;
          return acc;
        }, {});

        console.log('   Event types:', byType);
        console.log('   Last 5 events:');
        data.events.slice(-5).forEach(event => {
          const fromInfo = event.fromName ? ` from ${event.fromName}` : '';
          console.log(`     [${new Date(event.timestamp).toLocaleTimeString()}] ${event.type}${fromInfo}: ${event.message}`);
        });
        console.log('');
      });

      console.log('=== ACTIVITY BY EMAIL SENDER ===');
      Object.keys(bySender).forEach(sender => {
        const data = bySender[sender];
        console.log(`📧 Sender: ${data.name} (${sender})`);
        console.log(`   Events: ${data.events.length}`);
        console.log('');
      });

      // Log summary to activity log
      const mailboxCount = Object.keys(byMailbox).length;
      const senderCount = Object.keys(bySender).length;
      logActivity('success', `Activity Summary: ${history.length} events from ${senderCount} sender(s) across ${mailboxCount} mailbox(es)`);

      Object.keys(byMailbox).forEach(mailbox => {
        const data = byMailbox[mailbox];
        const mailboxShort = mailbox.split('@')[0];
        logActivity('info', `${mailboxShort}: ${data.events.length} events tracked`);
      });

      return { byMailbox, bySender };
    } catch (error) {
      console.error('Error showing mailbox activity:', error);
      logActivity('error', 'Failed to load mailbox activity');
    }
  }

  function clearActivityLog() {
    console.log('=== CLEAR LOG CLICKED ===');

    try {
      const activityLog = document.getElementById('activity-log');
      if (!activityLog) {
        console.error('activity-log element not found');
        return;
      }

      activityLog.innerHTML = '';
      logActivity('info', 'Activity log cleared (history preserved)');
      console.log('=== ACTIVITY LOG CLEARED ===');
      console.log('Timestamp:', new Date().toISOString());
      console.log('Note: Event history in localStorage is preserved');
    } catch (error) {
      console.error('Error clearing log:', error);
    }
  }

  function updateEventCounter() {
    try {
      const counterElement = document.getElementById('events-tracked');
      if (counterElement) {
        counterElement.textContent = eventCounter;
      } else {
        console.error('events-tracked element not found');
      }
    } catch (error) {
      console.error('Error updating event counter:', error);
    }
  }

  console.log('=== TASKPANE.JS FULLY LOADED (IIFE END) ===');
  console.log('Timestamp:', new Date().toISOString());

})();
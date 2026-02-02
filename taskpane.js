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

  // Initialize Office
  Office.onReady((info) => {
    console.log('=== TASKPANE OFFICE.ONREADY FIRED ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);

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

      console.log('DOM Elements Check:');
      console.log('  - clear-log:', clearLogBtn ? 'Found' : 'NOT FOUND');
      console.log('  - toggle-monitoring:', toggleMonitoringBtn ? 'Found' : 'NOT FOUND');
      console.log('  - trigger-test-event:', triggerTestBtn ? 'Found' : 'NOT FOUND');

      if (!clearLogBtn || !toggleMonitoringBtn || !triggerTestBtn) {
        throw new Error('Required DOM elements not found');
      }

      // Attach event handlers
      clearLogBtn.onclick = clearActivityLog;
      toggleMonitoringBtn.onclick = toggleMonitoring;
      triggerTestBtn.onclick = triggerTestEvent;

      console.log('Event handlers attached successfully');

      logActivity('info', 'InboxAgent taskpane initialized');

      // Check for event runtime
      checkEventRuntime();

      // Update current item
      updateCurrentItem();

      // Add Office event listeners
      addOfficeEventListeners();

      // Start deep monitoring immediately
      setTimeout(() => {
        startDeepMonitoring();
        logActivity('success', 'Deep monitoring started automatically');
      }, 500);

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

    logActivity('info', 'Test event triggered - check console for details');

    const item = Office.context.mailbox.item;
    if (item) {
      console.log('Current Item Details:');
      console.log('  - Item Type:', item.itemType);
      console.log('  - Item Class:', item.itemClass);
      console.log('  - Item ID:', item.itemId || 'No ID (new item)');
      console.log('  - Conversation ID:', item.conversationId);
      console.log('  - Read Status:', item.read);

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
          console.log('  - From:', JSON.stringify(value, null, 2));
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

    logActivity('info', 'Item changed - Loading new item details');
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
      return;
    }

    console.log('Item available, getting subject...');

    getPropertyValue(item, 'subject', (subject) => {
      const displaySubject = subject || '(No Subject)';
      currentItemElement.textContent =
        displaySubject.substring(0, 30) + (displaySubject.length > 30 ? '...' : '');

      console.log('=== CURRENT ITEM UPDATED ===');
      console.log('Subject:', displaySubject);
      console.log('Item Type:', item.itemType);
      console.log('Item ID:', item.itemId || 'New item (no ID)');
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

    console.log('=== MONITORING TOGGLED ===');
    console.log('Monitoring Active:', isMonitoring);
    console.log('Timestamp:', new Date().toISOString());
  }

  function startDeepMonitoring() {
    console.log('=== STARTING DEEP MONITORING ===');

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

    console.log('=== DEEP MONITORING STOPPED ===');
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
      logActivity('warning', `Email ${statusChange}`);

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

    const captureQueue = new Queue({ results: [], concurrency: 1 });
    const state = {
      itemType: item.itemType,
      itemId: item.itemId,
      itemClass: item.itemClass || null
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

    // Capture from
    captureQueue.push(cb => {
      getPropertyValue(item, 'from', (value) => {
        state.from = value;
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
      logActivity('error', `Email disappeared: "${previousItemState.subject}" - possibly marked as junk or deleted`);

      previousItemState = null;
      return;
    }

    if (!item) return;

    const checkQueue = new Queue({ results: [], concurrency: 1 });
    const currentState = {
      itemType: item.itemType,
      itemId: item.itemId,
      itemClass: item.itemClass || null
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
        currentState.from = value;
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

      // Check From change
      const oldFrom = JSON.stringify(oldState.from || null);
      const newFrom = JSON.stringify(newState.from || null);
      if (oldFrom !== newFrom) {
        const change = 'From address changed';
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

        let actionType = 'UNKNOWN';
        if (newState.subject && oldState.subject) {
          if (newState.subject.startsWith('RE:') || newState.subject.startsWith('Re:')) {
            actionType = 'REPLY';
          } else if (newState.subject.startsWith('FW:') || newState.subject.startsWith('Fw:')) {
            actionType = 'FORWARD';
          }
        }

        logActivity('success', `${actionType}: "${oldState.subject}"`);

        console.log('Detected Action:', actionType);
      }
    }
  }

  // Detect if email was marked as junk by monitoring item disappearance
  function detectJunkMarking(oldState, newState) {
    console.log('=== CHECKING FOR JUNK MARKING ===');

    // Case 1: Item ID changed but we're still in the same context
    if (oldState.itemId && newState.itemId && oldState.itemId !== newState.itemId) {
      console.log('Item ID changed - email may have been moved');
      logActivity('warning', 'Email moved or marked as junk/not junk');
    }

    // Case 2: Item became null (disappeared)
    if (oldState.itemId && !newState.itemId) {
      console.log('Item disappeared - likely moved to Junk or Deleted');
      logActivity('error', `Email disappeared: "${oldState.subject}" - possibly marked as junk`);

      console.log('=== EMAIL MARKED AS JUNK (LIKELY) ===');
      console.log('Subject:', oldState.subject);
      console.log('From:', oldState.from?.emailAddress);
      console.log('Item ID:', oldState.itemId);
    }
  }

  // Simplified folder change detection (no EWS needed)
  function detectFolderChangeSimple(oldState, newState) {
    console.log('=== CHECKING FOR FOLDER CHANGE (SIMPLE) ===');

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

          logActivity('warning', `Email may have been moved: "${newState.subject}"`);

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

      const activityItem = document.createElement('div');
      activityItem.className = `activity-item ${type}`;

      const time = document.createElement('div');
      time.className = 'activity-time';
      time.textContent = new Date().toLocaleTimeString();

      const msg = document.createElement('div');
      msg.className = 'activity-message';
      msg.textContent = message;

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
    } catch (error) {
      console.error('Error logging activity:', error);
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
      logActivity('info', 'Activity log cleared');
      console.log('=== ACTIVITY LOG CLEARED ===');
      console.log('Timestamp:', new Date().toISOString());
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
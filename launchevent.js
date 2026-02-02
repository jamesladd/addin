/* global Office */

console.log('=== LAUNCHEVENT.JS LOADING... ===');
console.log('Timestamp:', new Date().toISOString());
console.log('Window object available:', typeof window !== 'undefined');
console.log('Global object available:', typeof global !== 'undefined');

// IIFE to immediately attach all event listeners
(function () {
  'use strict';

  console.log('=== LAUNCHEVENT.JS IIFE EXECUTING ===');
  console.log('Timestamp:', new Date().toISOString());

  // Ensure Office is ready
  Office.onReady(() => {
    console.log('=== OFFICE READY IN LAUNCHEVENT ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Office.context available:', typeof Office.context !== 'undefined');
    console.log('Office.context.mailbox available:', typeof Office.context.mailbox !== 'undefined');
  });

  // Event Handlers

  function onNewMessageComposeHandler(event) {
    console.log('=== EVENT: OnNewMessageCompose ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));
    console.log('Event Type:', event.type);

    const item = Office.context.mailbox.item;
    if (item) {
      console.log('Item Type:', item.itemType);
      console.log('Item Class:', item.itemClass);

      item.subject.getAsync((result) => {
        console.log('Subject:', result.value);
      });
    }

    // Complete the event
    event.completed();
  }

  function onMessageAttachmentsChangedHandler(event) {
    console.log('=== EVENT: OnMessageAttachmentsChanged ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));

    const item = Office.context.mailbox.item;
    if (item && item.attachments) {
      console.log('Attachments Count:', item.attachments.length);
      console.log('Attachments:', JSON.stringify(item.attachments.map(a => ({
        id: a.id,
        name: a.name,
        size: a.size,
        type: a.attachmentType
      })), null, 2));
    }

    event.completed();
  }

  function onMessageRecipientsChangedHandler(event) {
    console.log('=== EVENT: OnMessageRecipientsChanged ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));

    const item = Office.context.mailbox.item;
    if (item) {
      if (item.to) {
        item.to.getAsync((result) => {
          console.log('To Recipients:', JSON.stringify(result.value, null, 2));
        });
      }
      if (item.cc) {
        item.cc.getAsync((result) => {
          console.log('CC Recipients:', JSON.stringify(result.value, null, 2));
        });
      }
      if (item.bcc) {
        item.bcc.getAsync((result) => {
          console.log('BCC Recipients:', JSON.stringify(result.value, null, 2));
        });
      }
    }

    event.completed();
  }

  function onInfoBarDismissClickedHandler(event) {
    console.log('=== EVENT: OnInfoBarDismissClicked ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));

    event.completed();
  }

  function onMessageSendHandler(event) {
    console.log('=== EVENT: OnMessageSend ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));

    const item = Office.context.mailbox.item;
    if (item) {
      item.subject.getAsync((subjectResult) => {
        console.log('Sending Message Subject:', subjectResult.value);

        item.to.getAsync((toResult) => {
          console.log('Sending To:', JSON.stringify(toResult.value, null, 2));

          if (item.attachments) {
            console.log('Attachments Count:', item.attachments.length);
          }

          // Allow send
          event.completed({ allowEvent: true });
        });
      });
    } else {
      event.completed({ allowEvent: true });
    }
  }

  function onMessageFromChangedHandler(event) {
    console.log('=== EVENT: OnMessageFromChanged ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));

    const item = Office.context.mailbox.item;
    if (item && item.from) {
      item.from.getAsync((result) => {
        console.log('New From Address:', JSON.stringify(result.value, null, 2));
      });
    }

    event.completed();
  }

  function onSensitivityLabelChangedHandler(event) {
    console.log('=== EVENT: OnSensitivityLabelChanged ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Object:', JSON.stringify(event, null, 2));

    const item = Office.context.mailbox.item;
    if (item) {
      console.log('Item Type:', item.itemType);
      // Sensitivity label information would be available in event object
    }

    event.completed();
  }

  console.log('=== REGISTERING EVENT HANDLERS ===');
  console.log('Timestamp:', new Date().toISOString());

  // Register functions globally for manifest
  if (typeof global !== 'undefined') {
    console.log('Registering handlers on global object');
    global.onNewMessageComposeHandler = onNewMessageComposeHandler;
    global.onMessageAttachmentsChangedHandler = onMessageAttachmentsChangedHandler;
    global.onMessageRecipientsChangedHandler = onMessageRecipientsChangedHandler;
    global.onInfoBarDismissClickedHandler = onInfoBarDismissClickedHandler;
    global.onMessageSendHandler = onMessageSendHandler;
    global.onMessageFromChangedHandler = onMessageFromChangedHandler;
    global.onSensitivityLabelChangedHandler = onSensitivityLabelChangedHandler;
  }

  // Also register on window for browser context
  if (typeof window !== 'undefined') {
    console.log('Registering handlers on window object');
    window.onNewMessageComposeHandler = onNewMessageComposeHandler;
    window.onMessageAttachmentsChangedHandler = onMessageAttachmentsChangedHandler;
    window.onMessageRecipientsChangedHandler = onMessageRecipientsChangedHandler;
    window.onInfoBarDismissClickedHandler = onInfoBarDismissClickedHandler;
    window.onMessageSendHandler = onMessageSendHandler;
    window.onMessageFromChangedHandler = onMessageFromChangedHandler;
    window.onSensitivityLabelChangedHandler = onSensitivityLabelChangedHandler;

    // Log that they are registered
    console.log('Window handlers registered:');
    console.log('  - onNewMessageComposeHandler:', typeof window.onNewMessageComposeHandler);
    console.log('  - onMessageAttachmentsChangedHandler:', typeof window.onMessageAttachmentsChangedHandler);
    console.log('  - onMessageRecipientsChangedHandler:', typeof window.onMessageRecipientsChangedHandler);
    console.log('  - onInfoBarDismissClickedHandler:', typeof window.onInfoBarDismissClickedHandler);
    console.log('  - onMessageSendHandler:', typeof window.onMessageSendHandler);
    console.log('  - onMessageFromChangedHandler:', typeof window.onMessageFromChangedHandler);
    console.log('  - onSensitivityLabelChangedHandler:', typeof window.onSensitivityLabelChangedHandler);
  }

  console.log('=== ALL EVENT HANDLERS REGISTERED ===');
  console.log('Registered Handlers:', [
    'onNewMessageComposeHandler',
    'onMessageAttachmentsChangedHandler',
    'onMessageRecipientsChangedHandler',
    'onInfoBarDismissClickedHandler',
    'onMessageSendHandler',
    'onMessageFromChangedHandler',
    'onSensitivityLabelChangedHandler'
  ]);

})();

console.log('=== LAUNCHEVENT.JS FULLY LOADED ===');
console.log('Timestamp:', new Date().toISOString());
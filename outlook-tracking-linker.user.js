// ==UserScript==
// @name         Outlook Web - Shipment Tracking Linker
// @namespace    github.com/ruffy314/
// @author       Ruffy314
// @version      1.0.2
// @description  Turn tracking numbers into links in Outlook Web
// @match        https://outlook.office.com/*
// @match        https://outlook.cloud.microsoft/*
// @match        https://outlook.cloud.microsoft/mail/*
// @updateURL    https://cdn.jsdelivr.net/gh/Ruffy314/outlook-web-shipment-tracking-linker@main/outlook-tracking-linker.user.js
// @downloadURL  https://cdn.jsdelivr.net/gh/Ruffy314/outlook-web-shipment-tracking-linker@main/outlook-tracking-linker.user.js
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // Define shipping configurations
  const SHIPPING_COMPANIES = [
    {
      name: 'UPS',
      regex: /\b1Z[0-9A-Z]{16}\b/gi,
      linkTemplate: (trackingNumber) => `https://www.ups.com/track?tracknum=${encodeURIComponent(trackingNumber)}`,
      active: true,
    },
    {
      name: 'DHL',
      regex: /\b\d{10}\b/gi, // Sample DHL tracking pattern (adjust if necessary)
      linkTemplate: (trackingNumber) => `https://www.dhl.com/en/express/tracking.html?AWB=${encodeURIComponent(trackingNumber)}&brand=DHL`,
      active: false,
    },
    {
      name: 'Amazon',
      regex: /\bTBA[A-Z0-9]{12}\b/gi, // Sample Amazon pattern (adjust if necessary)
      linkTemplate: (trackingNumber) => `https://www.amazon.com/progress-tracker/package/${encodeURIComponent(trackingNumber)}`,
      active: false,
    },
    {
      name: 'Time:Matters',
      regex: /\bS\d{6}\b/gi,
      linkTemplate: (trackingNumber) => `https://booking.time-matters.com/en-US/tracking/${encodeURIComponent(trackingNumber)}`,
      active: true,
    },
  ];

  // Tags to skip (don't linkify inside these)
  const SKIP_TAGS = new Set(['A', 'SCRIPT', 'STYLE', 'PRE', 'CODE', 'TEXTAREA']);

  function isInComposeArea(node) {
    let parent = node.parentElement;
    while (parent) {
      if (parent.contentEditable === 'true') return true;
      parent = parent.parentElement;
    }
    return false;
  }

  function hasSkipAncestor(node) {
    let parent = node.parentElement;
    while (parent) {
      if (SKIP_TAGS.has(parent.tagName)) return true;
      parent = parent.parentElement;
    }
    return false;
  }

  // Modify createLink to accept a linkTemplate
  function createLink(trackingNumber, linkTemplate) {
    const a = document.createElement('a');
    a.href = linkTemplate(trackingNumber);
    a.textContent = trackingNumber;
    a.target = '_blank';
    a.rel = 'noreferrer noopener';
    a.style.color = '#0078d4';
    return a;
  }

  function processTextNode(node) {
    if (!node || node.nodeType !== Node.TEXT_NODE) return;
    if (isInComposeArea(node)) return;
    if (hasSkipAncestor(node)) return;

    const text = node.nodeValue;
    const fragment = document.createDocumentFragment();
    let lastIndex = 0;
    let matchFound = false;

    // Loop through each shipping company and process matches
    SHIPPING_COMPANIES.filter(obj => obj?.active).forEach(({ regex, linkTemplate }) => {
      regex.lastIndex = 0; // Reset regex state
    let match;
      while ((match = regex.exec(text)) !== null) {
        matchFound = true;
      const matchText = match[0];
      const offset = match.index;

      // Append text before match
      if (offset > lastIndex) {
        fragment.appendChild(document.createTextNode(text.slice(lastIndex, offset)));
      }

        // Append link
        fragment.appendChild(createLink(matchText, linkTemplate));
      lastIndex = offset + matchText.length;
    }
    });
    if (!matchFound) return; // No replacements performed
    fragment.appendChild(document.createTextNode(text.slice(lastIndex)));
    node.replaceWith(fragment);
  }

  // Walk text nodes under root and process them
  function walkAndProcess(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
    let node;
    while ((node = walker.nextNode())) {
      processTextNode(node);
    }
  }

  // Full initial scan
  walkAndProcess(document.body);

  // MutationObserver: process added nodes only, debounce to avoid high-frequency re-scans
  let scheduled = null;
  const pendingRoots = new Set();

  function scheduleProcess() {
    if (scheduled) return;
    scheduled = setTimeout(() => {
      // Process each unique root previously added
      const roots = Array.from(pendingRoots);
      pendingRoots.clear();
      scheduled = null;
      // Use requestIdleCallback when available to reduce jank
      const runner = () => {
        for (const root of roots) {
          try {
            // If the added node is a text node, process it directly; otherwise walk its subtree
            if (root.nodeType === Node.TEXT_NODE) {
              processTextNode(root);
            } else if (root.nodeType === Node.ELEMENT_NODE) {
              walkAndProcess(root);
            } else {
              // fallback: walk document.body
              walkAndProcess(document.body);
            }
          } catch (e) {
            // swallow individual errors to keep observer working
            console.error('UPS linker processing error', e);
          }
        }
      };
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(runner, { timeout: 500 });
      } else {
        // run on next macrotask to batch mutations
        setTimeout(runner, 0);
      }
    }, 200); // debounce window (200ms)
  }

  const observer = new MutationObserver((records) => {
    for (const record of records) {
      // Prefer addedNodes to limit work
      if (record.addedNodes && record.addedNodes.length) {
        for (const node of record.addedNodes) {
          // Skip nodes that are clearly inside skip tags
          if (node.nodeType === Node.ELEMENT_NODE && SKIP_TAGS.has(node.tagName)) continue;
          pendingRoots.add(node);
        }
      } else {
        // If no addedNodes (attribute changes etc.), schedule a light re-scan of the target
        if (record.target) pendingRoots.add(record.target);
      }
    }
    scheduleProcess();
  });

  observer.observe(document.body, { childList: true, subtree: true });

})();


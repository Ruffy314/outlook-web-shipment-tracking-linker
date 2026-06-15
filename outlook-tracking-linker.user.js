// ==UserScript==
// @name         Outlook Web - Shipment Tracking Linker
// @namespace    github.com/ruffy314/
// @author       Ruffy314
// @version      1.0.5
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

    /**
   * @typedef {Object} ShippingCompany
   * @property {string} name - Human-readable carrier name.
   * @property {RegExp} regex - Pattern used to find tracking numbers for this carrier.
   * @property {(trackingNumber: string) => string} linkTemplate - Function returning the tracking URL for a given number.
   * @property {boolean} active - Whether matching for this carrier is enabled.
   */
  /** @type {ShippingCompany[]} */
  const SHIPPING_COMPANIES = [
    {
      name: 'UPS',
      regex: /\b1Z[0-9A-Z]{16}\b/gi,
      linkTemplate: (trackingNumber) => `https://www.ups.com/track?tracknum=${encodeURIComponent(trackingNumber)}`,
      active: true,
    },
    {
      name: 'UPS weird spaces',
      regex: /\b1Z\s+[0-9A-Z]{3}\s+[0-9A-Z]{3}\s+[0-9A-Z]{2}\s+[0-9A-Z]{4}\s+[0-9A-Z]{4}\b/gi,
      linkTemplate: (trackingNumber) => `https://www.ups.com/track?tracknum=${encodeURIComponent(trackingNumber.replaceAll(/\s+/g, ""))}`,
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
    {
      name: 'Time:Matters written by idiots',
      regex: /\b(S\d{6})[A-Z]+\b/gi,
      linkTemplate: (trackingNumber) => `https://booking.time-matters.com/en-US/tracking/${encodeURIComponent(trackingNumber.slice(0, 7))}`,
      active: true,
    },
  ];

    /** Set of tag names to skip when linkifying.
   * @type {Set<string>}
   */
  const SKIP_TAGS = new Set(['A', 'SCRIPT', 'STYLE', 'PRE', 'CODE', 'TEXTAREA']);

    /**
   * Determine whether a node is within a compose area (contentEditable) in Outlook.
   * @param {Node} node
   * @returns {boolean}
   */
  function isInComposeArea(node) {
    let parent = node.parentElement;
    while (parent) {
      if (parent.contentEditable === 'true') return true;
      parent = parent.parentElement;
    }
    return false;
  }

    /**
   * Check if the node has an ancestor with a tag that should be skipped.
   * @param {Node} node
   * @returns {boolean}
   */
  function hasSkipAncestor(node) {
    let parent = node.parentElement;
    while (parent) {
      if (SKIP_TAGS.has(parent.tagName)) return true;
      parent = parent.parentElement;
    }
    return false;
  }

    // Modify createLink to accept a linkTemplate
  /**
   * Create an anchor for a tracking number using the provided link template.
   * @param {string} trackingNumber
   * @param {(trackingNumber: string) => string} linkTemplate
   * @returns {HTMLAnchorElement}
   */
  function createLink(trackingNumber, linkTemplate) {
    const a = document.createElement('a');
    a.href = linkTemplate(trackingNumber);
    a.textContent = trackingNumber;
    a.target = '_blank';
    a.rel = 'noreferrer noopener';
    a.style.color = '#0078d4';
    return a;
  }

    /**
   * Linkify tracking numbers within a single text node.
   * Collects matches from all active carriers, resolves overlaps, and replaces the node.
   * Safely no-ops inside compose areas and skip-tag ancestors.
   * @param {Node} node
   * @returns {void}
   */
  function processTextNode(node) {
    if (!node || node.nodeType !== Node.TEXT_NODE) return;
    if (isInComposeArea(node)) return;
    if (hasSkipAncestor(node)) return;

    const text = node.nodeValue;
    /** @type {Array<{start: number, end: number, text: string, linkTemplate: (trackingNumber: string) => string}>} */
    const matches = [];

    for (const { regex, linkTemplate, active } of SHIPPING_COMPANIES) {
      if (!active) continue;
      // Reset regex state (global regexes keep state)
      regex.lastIndex = 0;
      let m;
      while ((m = regex.exec(text)) !== null) {
        matches.push({
          start: m.index,
          end: m.index + m[0].length,
          text: m[0],
          linkTemplate,
        });
        // Safety: avoid zero-length infinite loop
        if (m.index === regex.lastIndex) regex.lastIndex++;
      }
    }

    if (matches.length === 0) return;

    // Sort by start index, then prefer longer matches to minimize overlaps
    matches.sort((a, b) => a.start - b.start || (b.end - b.start) - (a.end - a.start));

    // Build fragment while skipping overlaps
    const fragment = document.createDocumentFragment();
    let lastIndex = 0;

    for (const match of matches) {
      if (match.start < lastIndex) {
        // Overlaps with a previously accepted match: skip
        continue;
      }
      if (match.start > lastIndex) {
        fragment.appendChild(document.createTextNode(text.slice(lastIndex, match.start)));
      }
      fragment.appendChild(createLink(match.text, match.linkTemplate));
      lastIndex = match.end;
    }
    if (lastIndex < text.length) {
      fragment.appendChild(document.createTextNode(text.slice(lastIndex)));
    }

    node.replaceWith(fragment);
  }

    // Walk text nodes under root and process them
  /**
   * Collect text nodes beneath a root and process them after collection to avoid walker invalidation.
   * @param {Node} root
   * @returns {void}
   */
  function walkAndProcess(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
    /** @type {Node[]} */
    const nodes = [];
    let node;
    while ((node = walker.nextNode())) {
      nodes.push(node);
    }
    for (const n of nodes) {
      processTextNode(n);
    }
  }

  // Full initial scan
  walkAndProcess(document.body);

    // MutationObserver: process added nodes only, debounce to avoid high-frequency re-scans
  /** Handle for the debounce timer. */
  /** @type {number | null} */
  let scheduled = null;
  /** Set of nodes queued for processing. */
  /** @type {Set<Node>} */
  const pendingRoots = new Set();

    /**
   * Debounced processing scheduler for pending roots queued by the MutationObserver.
   * Batches work and runs it on the next macrotask.
   * @returns {void}
   */
  function scheduleProcess() {
    if (scheduled) return;
    scheduled = setTimeout(() => {
      const roots = Array.from(pendingRoots);
      pendingRoots.clear();
      scheduled = null;

      const runner = () => {
        for (const root of roots) {
          try {
            if (root.nodeType === Node.TEXT_NODE) {
              // Skip text nodes inside skip tags early
              const parent = root.parentElement;
              if (parent && SKIP_TAGS.has(parent.tagName)) continue;
              processTextNode(root);
            } else if (root.nodeType === Node.ELEMENT_NODE) {
              walkAndProcess(root);
            }
          } catch (e) {
            console.error('UPS linker processing error', e);
          }
        }
      };

      // Run on next macrotask to batch mutations without waiting for "idle"
      setTimeout(runner, 0);
    }, 100); // shorter debounce for better responsiveness
  }

  /** @type {MutationObserver} */
  const observer = new MutationObserver((/** @type {MutationRecord[]} */ records) => {
    for (const record of records) {
      if (record.addedNodes && record.addedNodes.length) {
        for (const node of record.addedNodes) {
          // Skip elements that we never want to process
          if (node.nodeType === Node.ELEMENT_NODE) {
            if (SKIP_TAGS.has(node.tagName)) continue;
            pendingRoots.add(node);
          } else if (node.nodeType === Node.TEXT_NODE) {
            // Skip text nodes already inside skip tags (e.g., text inside <a>)
            const parent = node.parentElement;
            if (parent && SKIP_TAGS.has(parent.tagName)) continue;
            pendingRoots.add(node);
          }
        }
      } else if (record.target) {
        // Attribute changes etc. Lightly rescan if the target isn't in a skip tag
        const t = record.target;
        if (t.nodeType === Node.ELEMENT_NODE) {
          if (!SKIP_TAGS.has(t.tagName)) pendingRoots.add(t);
        } else if (t.nodeType === Node.TEXT_NODE) {
          const parent = t.parentElement;
          if (!parent || !SKIP_TAGS.has(parent.tagName)) pendingRoots.add(t);
        }
      }
    }
    scheduleProcess();
  });

  observer.observe(document.body, { childList: true, subtree: true });

})();


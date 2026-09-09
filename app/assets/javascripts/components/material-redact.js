// app/assets/javascripts/components/material-redact.js

(() => {
  var viewer = document.getElementById('material-viewer')
  if (!viewer) return

  // ?redactHandoff=1 (see case--material.js) reinstates "Redact this
  // document" as a handoff to an external redaction tool, opened in a new
  // tab, instead of the in-page drag-select popover (material-redact-
  // popover.js — not even loaded on this page in handoff mode, see
  // material-panel.njk). Default (no flag): no standalone button at all —
  // drag-selecting text triggers the popover unconditionally instead.
  // ?tab=3 lands directly on that app's "Review and redact" tab — its own
  // housekeeping.js reads a `tab` URL parameter on load and calls
  // showTabByNumber(3), the same number its own click handler derives
  // from the "tab-3-content" class on that tab's link.
  var HANDOFF_URL = 'https://casework-housekeeping-merge-85fd6f0cf819.herokuapp.com/version-1-3/A-index?tab=3'
  var isHandoff = !!window.__DCF_REDACT_HANDOFF__

  // Standalone toolbar button — commented out of the injection call below
  // (not deleted) so it can be flipped back on quickly for a stakeholder
  // demo/comparison. "Redact" inside the Document actions menu (see
  // injectRedactMenuItem) is the live entry point for now.
  function injectRedactHandoffLink (toolbarRight) {
    if (toolbarRight.querySelector('[data-action="redact-document"]')) return
    var docActions = toolbarRight.querySelector('details[data-menu="document"]')
    if (!docActions) return

    var link = document.createElement('a')
    link.href = HANDOFF_URL
    link.target = '_blank'
    link.rel = 'noopener'
    link.className = 'govuk-button govuk-button--secondary govuk-!-margin-bottom-0 govuk-!-margin-right-3'
    link.setAttribute('data-action', 'redact-document')
    link.textContent = 'Redact this document'

    toolbarRight.insertBefore(link, docActions)
  }

  // Same handoff as injectRedactHandoffLink above, just as the first item
  // in the Document actions menu instead of a separate toolbar button —
  // shares its data-action="redact-document", so the existing click
  // handler below already covers it with no extra wiring.
  function injectRedactMenuItem (toolbarRight) {
    var list = toolbarRight.querySelector('.dcf-action-menu__list')
    if (!list || list.querySelector('[data-action="redact-document"]')) return

    var item = document.createElement('li')
    item.className = 'dcf-action-menu__item'
    item.innerHTML = '<a href="' + HANDOFF_URL + '" target="_blank" rel="noopener" class="govuk-link dcf-action-menu__link" data-action="redact-document">Redact this document</a>'
    list.insertBefore(item, list.firstChild)
  }

  function injectAiRedactionMenuItem (toolbarRight) {
    var list = toolbarRight.querySelector('.dcf-action-menu__list')
    if (!list || list.querySelector('[data-action="ai-redact-document"]')) return

    var item = document.createElement('li')
    item.className = 'dcf-action-menu__item'
    item.innerHTML = '<a href="#" class="govuk-link dcf-action-menu__link" data-action="ai-redact-document">AI redaction</a>'
    list.appendChild(item)
  }

  // Aligns the dropdown to the right edge of the "Document actions"
  // button instead of the left (see .dcf-action-menu--right in
  // _dcf-moj-button-menu-fallback.scss) — set here rather than on the
  // <details> markup itself, which is built in material-viewer.js.
  function rightAlignDocActionsMenu (toolbarRight) {
    var docActions = toolbarRight.querySelector('details[data-menu="document"]')
    if (docActions) docActions.classList.add('dcf-action-menu--right')
  }

  var observer = new MutationObserver(function () {
    var toolbarRight = viewer.querySelector('.dcf-viewer__toolbar-right')
    if (toolbarRight) {
      // if (isHandoff) injectRedactHandoffLink(toolbarRight)
      if (isHandoff) injectRedactMenuItem(toolbarRight)
      injectAiRedactionMenuItem(toolbarRight)
      rightAlignDocActionsMenu(toolbarRight)
    }
  })
  observer.observe(viewer, { childList: true, subtree: true })

  viewer.addEventListener('click', function (e) {
    // material-viewer.js's own delegated click handler matches any
    // [data-action] element and calls e.preventDefault() unconditionally
    // before checking which action it is, then does nothing for actions
    // it doesn't recognise (both of these are exactly that) — so the
    // anchors' native href/target behaviour never fires. Doing the
    // navigation explicitly in JS here works regardless, since an
    // earlier preventDefault() on the event doesn't stop other listeners
    // from running.
    var handoffLink = e.target.closest('[data-action="redact-document"]')
    if (handoffLink) {
      window.open(HANDOFF_URL, '_blank', 'noopener')
      return
    }

    var aiLink = e.target.closest('[data-action="ai-redact-document"]')
    if (!aiLink) return

    var activeTab = viewer.querySelector('.dcf-doc-tab.is-active')
    if (!activeTab) return

    var url    = activeTab.getAttribute('data-url')     || ''
    var title  = activeTab.getAttribute('data-title')   || 'Document'
    var itemId = activeTab.getAttribute('data-item-id') || ''
    var caseId = window.location.pathname.split('/')[2]

    window.location.href =
      '/cases/' + caseId + '/material/redact/scan' +
      '?url='    + encodeURIComponent(url) +
      '&title='  + encodeURIComponent(title) +
      '&itemId=' + encodeURIComponent(itemId)
  })
})()

(() => {
  var viewer = document.getElementById('material-viewer')
  if (!viewer) return

  // "Redact this document" stays exactly where it was (a standalone
  // button next to the Document actions menu), but now launches the
  // in-place redact popover (material-redact-popover.js) instead of
  // navigating to the separate AI-assisted flow. That flow is still
  // reachable — moved into Document actions as "AI redaction" — since it
  // may be revisited in future.
  function injectRedactLink (toolbarRight) {
    if (toolbarRight.querySelector('[data-action="redact-document"]')) return
    var docActions = toolbarRight.querySelector('details[data-menu="document"]')
    if (!docActions) return

    var link = document.createElement('a')
    link.href = '#'
    link.className = 'govuk-button govuk-button--secondary govuk-!-margin-bottom-0 govuk-!-margin-right-3'
    link.setAttribute('data-action', 'redact-document')
    link.setAttribute('aria-pressed', 'false')
    link.textContent = 'Redact this document'

    toolbarRight.insertBefore(link, docActions)
  }

  function injectAiRedactionMenuItem (toolbarRight) {
    var list = toolbarRight.querySelector('.dcf-action-menu__list')
    if (!list || list.querySelector('[data-action="ai-redact-document"]')) return

    var item = document.createElement('li')
    item.className = 'dcf-action-menu__item'
    item.innerHTML = '<a href="#" class="govuk-link dcf-action-menu__link" data-action="ai-redact-document">AI redaction</a>'
    list.appendChild(item)
  }

  var observer = new MutationObserver(function () {
    var toolbarRight = viewer.querySelector('.dcf-viewer__toolbar-right')
    if (toolbarRight) {
      injectRedactLink(toolbarRight)
      injectAiRedactionMenuItem(toolbarRight)
    }
  })
  observer.observe(viewer, { childList: true, subtree: true })

  viewer.addEventListener('click', function (e) {
    var aiLink = e.target.closest('[data-action="ai-redact-document"]')
    if (aiLink) {
      e.preventDefault()

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
      return
    }

    var redactLink = e.target.closest('[data-action="redact-document"]')
    if (!redactLink) return
    e.preventDefault()

    if (window.DCFMaterialRedactPopover) {
      window.DCFMaterialRedactPopover.toggle(redactLink)
    }
  })
})()

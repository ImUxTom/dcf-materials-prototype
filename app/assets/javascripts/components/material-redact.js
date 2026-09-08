(() => {
  var viewer = document.getElementById('material-viewer')
  if (!viewer) return

  // The standalone "Redact this document" button (and the popover it used
  // to gate) has been removed for now — drag-selecting text in the
  // document triggers the redact popover unconditionally instead (see
  // material-redact-popover.js), matching how the factual summary's
  // redact popover has no separate "enter redact mode" step either.
  // Reserved for reintroduction in a future version as the trigger for an
  // external redaction tool, opened in a new tab, rather than this
  // in-page popover.
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
    if (toolbarRight) injectAiRedactionMenuItem(toolbarRight)
  })
  observer.observe(viewer, { childList: true, subtree: true })

  viewer.addEventListener('click', function (e) {
    var aiLink = e.target.closest('[data-action="ai-redact-document"]')
    if (!aiLink) return
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
  })
})()

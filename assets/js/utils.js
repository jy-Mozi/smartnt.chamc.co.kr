'use strict';

/**
 * Form
 */
// 폼 데이터 확인용
export const logFormData = formData => {
  console.group('FormData');
  if (formData instanceof FormData) {
    for (let pair of formData.entries()) {
      console.log(pair[0] + ': ' + pair[1]);
    }
  } else {
    console.log(formData);
  }
  console.groupEnd('FormData');
};

// 폼 인풋 값 초기화
export const resetFormValues = form => {
  const inputs = form.querySelectorAll('input, textarea');
  const selects = form.querySelectorAll('select');

  if (inputs) inputs.forEach(input => (input.value = ''));
  if (selects) selects.forEach(select => (select.selectedIndex = 0));
};

/**
 * Modal
 */
// 모달 토글 메소드
export const modalMethods = {
  open(modal) {
    modal.classList.add('open');
  },
  close(modal) {
    const form = modal.querySelector('form');

    modal.classList.remove('open');
    form && resetFormValues(form);
  }
};

/**
 * DatePicker
 */
const datepickerOptions = {
  language: 'ko',
  autohide: true,
  orientation: 'left'
};

// 날짜 선택
export const configDatePicker = datePicker => {
  return new Datepicker(datePicker, datepickerOptions);
};

// 기간 선택
export const configRangePicker = rangePicker => {
  return new DateRangePicker(rangePicker, datepickerOptions);
};

/**
 * CKEditor
 */
export const editorOptions = {
  // https://ckeditor.com/docs/ckeditor5/latest/features/toolbar/toolbar.html#extended-toolbar-configuration-format
  toolbar: {
    items: [
      'undo',
      'redo',
      '|',
      'heading',
      '|',
      'fontSize',
      'fontFamily',
      'fontColor',
      'fontBackgroundColor',
      '|',
      'bold',
      'italic',
      'strikethrough',
      'underline',
      '|',
      'uploadImage',
      'mediaEmbed',
      'insertTable',
      '|',
      'link',
      'blockQuote',
      'horizontalLine',
      'code',
      'codeBlock',
      '|',
      'numberedList',
      'bulletedList',
      '|',
      'outdent',
      'indent',
      'alignment',
      '|',
      'sourceEditing'
    ],
    shouldNotGroupWhenFull: false
  },
  language: 'ko',
  list: {
    properties: {
      styles: true,
      startIndex: true,
      reversed: true
    }
  },
  // https://ckeditor.com/docs/ckeditor5/latest/features/headings.html#configuration
  heading: {
    options: [
      { model: 'paragraph', title: 'Paragraph', class: 'ck-heading_paragraph' },
      { model: 'heading1', view: 'h1', title: 'Heading 1', class: 'ck-heading_heading1' },
      { model: 'heading2', view: 'h2', title: 'Heading 2', class: 'ck-heading_heading2' },
      { model: 'heading3', view: 'h3', title: 'Heading 3', class: 'ck-heading_heading3' },
      { model: 'heading4', view: 'h4', title: 'Heading 4', class: 'ck-heading_heading4' },
      { model: 'heading5', view: 'h5', title: 'Heading 5', class: 'ck-heading_heading5' },
      { model: 'heading6', view: 'h6', title: 'Heading 6', class: 'ck-heading_heading6' }
    ]
  },
  // https://ckeditor.com/docs/ckeditor5/latest/features/editor-placeholder.html#using-the-editor-configuration
  placeholder: '내용을 입력해주세요.',
  // https://ckeditor.com/docs/ckeditor5/latest/features/font.html#configuring-the-font-family-feature
  fontFamily: {
    options: ['default', 'Pretendard, sans-serif'],
    supportAllValues: true
  },
  // https://ckeditor.com/docs/ckeditor5/latest/features/font.html#configuring-the-font-size-feature
  fontSize: {
    options: [10, 12, 14, 'default', 18, 20, 22],
    supportAllValues: true
  },
  // https://ckeditor.com/docs/ckeditor5/latest/features/general-html-support.html#enabling-all-html-features
  htmlSupport: {
    allow: [
      {
        name: /.*/,
        attributes: true,
        classes: true,
        styles: true
      }
    ]
  },
  // https://ckeditor.com/docs/ckeditor5/latest/features/html-embed.html#content-previews
  htmlEmbed: {
    showPreviews: false
  },
  // https://ckeditor.com/docs/ckeditor5/latest/features/link.html#custom-link-attributes-decorators
  link: {
    decorators: {
      addTargetToExternalLinks: true,
      defaultProtocol: 'https://',
      toggleDownloadable: {
        mode: 'manual',
        label: 'Downloadable',
        attributes: {
          download: 'file'
        }
      }
    }
  },
  removePlugins: [
    'ExportPdf',
    'ExportWord',
    'AIAssistant',
    'CKBox',
    'CKFinder',
    'EasyImage',
    // https://ckeditor.com/docs/ckeditor5/latest/features/images/image-upload/base64-upload-adapter.html
    // https://ckeditor.com/docs/ckeditor5/latest/features/images/image-upload/image-upload.html
    // 'Base64UploadAdapter',
    'MultiLevelList',
    'RealTimeCollaborativeComments',
    'RealTimeCollaborativeTrackChanges',
    'RealTimeCollaborativeRevisionHistory',
    'PresenceList',
    'Comments',
    'TrackChanges',
    'TrackChangesData',
    'RevisionHistory',
    'Pagination',
    'WProofreader',
    'MathType',
    'SlashCommand',
    'Template',
    'DocumentOutline',
    'FormatPainter',
    'TableOfContents',
    'PasteFromOfficeEnhanced',
    'CaseChange'
  ]
};

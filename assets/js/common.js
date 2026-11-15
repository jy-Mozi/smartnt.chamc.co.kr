'use strict';

import { logFormData, resetFormValues, modalMethods, configDatePicker, configRangePicker, editorOptions } from '/assets/js/utils.js';

{
  const elements = {
    // form
    forms: document.querySelectorAll('form'),
    // gnb
    gnbButtons: document.querySelectorAll('.gnb-button'),
    // setting
    setting: document.getElementById('setting'),
    // tab
    tabButtons: document.querySelectorAll('.tab-button'),
    // modal
    modalButtons: document.querySelectorAll('.modal-button'),
    // file
    fileInputs: document.querySelectorAll('.file-input'),
    // date picker
    datePickers: document.querySelectorAll('.date-picker input'),
    rangePickers: document.querySelectorAll('.range-picker'),
    // time picker
    timePickers: document.querySelectorAll('.time-picker'),
    // editor
    editors: document.querySelectorAll('.editor'),
    calendars: document.querySelectorAll('.fullcalendar')
  };
  const data = {
    calanderEvents: [
      {
        title: '주간',
        date: '2024-08-29',
        display: 'block',
        classNames: ['event-week']
      },
      {
        title: '주휴',
        date: '2024-08-28',
        display: 'list-item',
        classNames: ['event-weekly-holiday']
      },
      {
        title: '휴일',
        date: '2024-08-27',
        display: 'block',
        classNames: ['event-holiday']
      },
      {
        title: '연차',
        date: '2024-08-25',
        display: 'block',
        classNames: ['event-annual']
      },
      {
        title: '연차승인',
        date: '2024-08-25',
        display: 'list-item',
        classNames: ['event-annual-approval']
      },
      {
        title: '연차',
        date: '2024-08-26',
        display: 'block',
        classNames: ['event-annual']
      },
      {
        title: '연차거절',
        date: '2024-08-26',
        display: 'list-item',
        classNames: ['event-annual-rejection']
      },
      {
        title: '종합검진 (3/50)',
        date: '2024-08-01',
        display: 'block'
      },
      {
        title: '대장 (0/5)',
        date: '2024-08-01',
        display: 'block'
      },
      {
        title: '심초(0/3)',
        date: '2024-08-01',
        display: 'block'
      },
      {
        title: '유방 (0/1)',
        date: '2024-08-01',
        display: 'block'
      },
      {
        title: 'MRI (0/1)',
        date: '2024-08-01',
        display: 'block'
      },
      {
        title: 'CT (3/30)',
        date: '2024-08-01',
        display: 'block'
      }
    ],
    smallCalendarEvents: [
      {
        title: '반(오전)',
        date: '2024-08-31',
        display: 'block'
      },
      {
        title: '반(오후)',
        date: '2024-08-02',
        display: 'block'
      }
    ]
  };
  const methods = {
    // date picker
    setDatePickers() {
      elements.datePickers.forEach(datePicker => configDatePicker(datePicker));
    },
    setRangePickers() {
      elements.rangePickers.forEach(rangePicker => configRangePicker(rangePicker));
    },
    // time picker
    populatePicker(picker, values) {
      values.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.text = value;
        picker.appendChild(option);
      });
    },
    setTimePickers() {
      elements.timePickers.forEach(timePicker => {
        const hourPicker = timePicker.querySelector('.hour-picker');
        const minutePicker = timePicker.querySelector('.minute-picker');

        const hours = Array.from({ length: 24 }, (_, index) => String(index).padStart(2, '0'));
        const minutes = Array.from({ length: 60 }, (_, index) => String(index).padStart(2, '0'));

        this.populatePicker(hourPicker, hours);
        this.populatePicker(minutePicker, minutes);
      });
    },
    // editor
    createEditor() {
      elements.editors.forEach(editor => {
        CKEDITOR.ClassicEditor.create(editor, { ...editorOptions })
          .then(newEditor => (editor.ckeditor = newEditor))
          .catch(error => console.error(error));
      });
    },
    createCalendar() {
      elements.calendars.forEach(calendar => {
        const isSmallCalendar = calendar.classList.contains('small-calendar');

        const calendarInstance = new FullCalendar.Calendar(calendar, {
          ...(isSmallCalendar
            ? {}
            : {
                customButtons: {
                  monthlyBatchButton: {
                    text: '월별 일괄 등록하기',
                    click() {
                      alert('월별 일괄 등록하기!');
                    }
                  },
                  holidayButton: {
                    text: '공휴일설정',
                    click() {
                      alert('공휴일설정!');
                    }
                  }
                }
              }),
          dayCellContent: function (arg) {
            return (arg.dayNumberText = arg.dayNumberText.replace('일', ''));
          },
          headerToolbar: {
            center: 'prev title next',
            start: isSmallCalendar ? 'prev' : '',
            end: isSmallCalendar ? 'next' : 'monthlyBatchButton,holidayButton'
          },
          initialView: 'dayGridMonth',
          locale: 'ko',
          contentHeight: isSmallCalendar ? 'auto' : 700,
          eventOrder: 'sortIdx',
          events: isSmallCalendar ? data.smallCalendarEvents : data.calanderEvents
        });
        calendarInstance.render();
      });
    }
  };
  const handler = {
    // form
    submitForm(e) {
      e.preventDefault();

      const formData = new FormData(e.target);
      const editors = e.target.querySelectorAll('.editor');

      editors.forEach(editor => {
        const key = editor.getAttribute('name');
        const value = editor.ckeditor.getData();
        formData.append(key, value);
      });

//      logFormData(formData);
//      resetFormValues(e.target);
    },
    // gnb
    toggleLnb({ currentTarget }) {
      currentTarget.classList.toggle('active');
    },
    // setting
    toggleSetting({ target }) {
      if (!target.closest('button')) return;
      elements.setting.classList.toggle('open');
    },
    // tab
    setCurrentTab({ currentTarget }) {
      const currentTabButton = currentTarget.closest('.tab-list').querySelector('.tab-button.current');
      const currentTabContent = currentTarget.closest('.tab').querySelector('.tab-content.current');

      currentTabButton.classList.remove('current');
      currentTabContent.classList.remove('current');

      const tabId = currentTarget.dataset.id;
      const targetTabContent = document.querySelector(`.tab-content[data-id=${tabId}]`);

      currentTarget.classList.add('current');
      targetTabContent.classList.add('current');
    },
    // modal
    toggleModal({ currentTarget }) {
      const type = currentTarget.dataset.type;
      const modalId = currentTarget.dataset.id;
      const modal = document.querySelector(`.modal[data-id=${modalId}]`);

      modalMethods[type](modal);
    },
    // file
    changeFile({ target }) {
      const preview = target.closest('.file-input').querySelector('.preview');
      const fileName = target.nextElementSibling.querySelector('.file-name');
      const file = target.files[0];

      fileName.innerText = file.name ?? '등록된 파일이 없습니다.';
      if (preview && file) {
        const reader = new FileReader();
        reader.onload = e => (preview.src = e.target.result);
        reader.readAsDataURL(file);
      }
    }
  };
  const initialize = () => {
    // form
    elements.forms.forEach(form => form.addEventListener('submit', handler.submitForm));
    // gnb
    elements.gnbButtons.forEach(gnbButton => gnbButton.addEventListener('click', handler.toggleLnb));
    // tab
    elements.tabButtons.forEach(tabButton => tabButton.addEventListener('click', handler.setCurrentTab));
    // modal
    elements.modalButtons.forEach(modalButton => modalButton.addEventListener('click', handler.toggleModal));
    // file
    elements.fileInputs.forEach(fileInput => fileInput.addEventListener('input', handler.changeFile));
    // date picker
    methods.setDatePickers();
    methods.setRangePickers();
    // time picker
    methods.setTimePickers();
    // editor
    methods.createEditor();
    // calendar
    methods.createCalendar();
  };
  initialize();
}

import request from '@/utils/request'

export function getList(params) {
  return request({
    url: '/Persons',
    method: 'get',
    params
  })
}

export function putSendEmail(guid) {
  return request({
    url: '/SendEmail',
    method: 'put',
    data: {
      guid
    }
  })
}

export function putUpdate(data) {
  return request({
    url: '/Update',
    method: 'put',
    data: data
  })
}

export function putSendEmail_Regitster(guid) {
  return request({
    url: '/SendEmailRegitster',
    method: 'put',
    data: {
      guid
    }
  })
}

export function putSendEmail_Notification(guid) {
  return request({
    url: '/SendEmailNotification',
    method: 'put',
    data: {
      guid
    }
  })
}


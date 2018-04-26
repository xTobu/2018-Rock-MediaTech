import request from '@/utils/request'

export function login(username, password) {
  return request({
    url: '/JWT/Login',
    method: 'post',
    data: {
      username,
      password
    }
  })
}

export function getInfo(token) {
  return request({
    url: '/JWT/Authorization',
    method: 'get',
    data: {
      token
    }
  })
}

export function logout() {
  return request({
    url: '/JWT/Logout',
    method: 'post'
  })
}

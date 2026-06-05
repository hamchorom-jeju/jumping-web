const axios = require('axios');

export default async function handler(req, res) {
  const { action } = req.query;
  const GAS_URL = "https://script.google.com/macros/s/AKfycbwxF9RGlmGA0Mkbm-IfxBj17roEmpAnkxu75QCMycsx7H7y7SbrCimjN7UVxBxm0PHR1Q/exec";

  // OPTIONS 요청 대응 (CORS)
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  try {
    // 구글은 리다이렉트가 잦으므로 axios를 통해 데이터가 포함된 채로 끝까지 추적합니다.
    const response = await axios({
      method: 'post',
      url: GAS_URL,
      params: { action },
      data: req.body,
      timeout: 30000, // 구글 응답이 느릴 수 있으므로 30초 대기
      maxRedirects: 5
    });

    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.status(200).json(response.data);
  } catch (error) {
    console.error('GAS Proxy Error:', error.message);
    return res.status(500).json({ 
      error: '구글 서버 응답 오류', 
      details: error.message,
      hint: '구글 앱스 스크립트 배포가 [나(Me)] 기준, [모든 사람(Anyone)] 접근 가능으로 되어 있는지 확인해 주세요.'
    });
  }
}

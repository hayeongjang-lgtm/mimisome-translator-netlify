// Netlify Serverless Function
// 프론트엔드에서 받은 번역 요청을 Anthropic API로 중계
// API 키는 Netlify 환경변수에 저장되어 브라우저로 노출되지 않음

export const handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method not allowed' }),
    };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({
        error: 'ANTHROPIC_API_KEY 환경변수가 Netlify에 설정되지 않았습니다.',
      }),
    };
  }

  try {
    const body = typeof event.body === 'string' ? event.body : JSON.stringify(event.body);

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body,
    });

    const data = await response.json();

    return {
      statusCode: response.status,
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(data),
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ error: 'Anthropic API 호출 실패: ' + e.message }),
    };
  }
};

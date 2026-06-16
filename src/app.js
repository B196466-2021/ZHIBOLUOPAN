const mentors = [
  {
    name: '陈教授',
    school: '华东师范大学',
    field: '教育技术 / AI学习分析',
    match: 94,
    status: '招募中',
    tier: '稳申',
    reason: '研究方向与生成式AI辅助教学高度重合，近两年持续招收申请制博士。',
  },
  {
    name: '刘教授',
    school: '浙江大学',
    field: '数字治理 / 数据智能',
    match: 88,
    status: '材料审核中',
    tier: '冲刺',
    reason: '数据智能方向接近，但论文成果要求较高，适合作为冲刺目标。',
  },
  {
    name: '王教授',
    school: '南京大学',
    field: '社会计算 / 平台经济',
    match: 83,
    status: '少量名额',
    tier: '保底',
    reason: '交叉方向友好，适合作为跨学科申请的稳妥选择。',
  },
];

const groups = [
  ['冲刺导师', '目标略高，需要强化论文与研究计划。'],
  ['稳申导师', '背景与方向最匹配，建议优先陶瓷。'],
  ['保底导师', '录取风险较低，适合完善申请组合。'],
];

function renderMentorResults() {
  const wrap = document.querySelector('#mentorResults');
  wrap.innerHTML = mentors.map((mentor) => `
    <article class="mentor-result">
      <div>
        <strong>${mentor.name}</strong>
        <span>${mentor.school}</span>
        <p>${mentor.field}</p>
        <small>${mentor.reason}</small>
      </div>
      <div class="mentor-score">
        <b>${mentor.match}%</b>
        <em>${mentor.status}</em>
      </div>
    </article>
  `).join('');
}

function renderMatchColumns() {
  const wrap = document.querySelector('#matchColumns');
  wrap.innerHTML = groups.map(([title, desc], index) => {
    const mentor = mentors[index];
    return `
      <article class="match-column">
        <span>${title}</span>
        <h3>${mentor.name}</h3>
        <p>${desc}</p>
        <strong>${mentor.school}</strong>
        <em>${mentor.match}% 匹配</em>
      </article>
    `;
  }).join('');
}

document.querySelector('#searchButton')?.addEventListener('click', () => {
  renderMentorResults();
  document.querySelector('#mentorResults').classList.add('pulse');
  window.setTimeout(() => document.querySelector('#mentorResults').classList.remove('pulse'), 450);
});

document.querySelector('#refreshMatch')?.addEventListener('click', () => {
  mentors.push(mentors.shift());
  renderMatchColumns();
});

renderMentorResults();
renderMatchColumns();

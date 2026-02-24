const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat,
  TabStopType, TabStopPosition,
} = require('docx');
const fs = require('fs');

// IDEA Light Theme palette (ATS-friendly: light bg, dark text)
const C = {
  black:    '1A1A1A',
  keyword:  'B06000',  // orange-brown — keywords
  type:     '00627A',  // teal — types/section headers
  string:   '067D17',  // green — strings/companies
  comment:  '8C8C8C',  // gray — secondary info
  purple:   '7832A8',  // purple — annotations/labels
  link:     '2470B3',  // blue — links
  white:    'FFFFFF',
  headerBg: 'F8F8F8',
  divider:  'E0E0E0',
  tagBg:    'F0F4FF',
};

const noBorder = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helpers ──────────────────────────────────────────────────────────────────

function t(text, color = C.black, opts = {}) {
  return new TextRun({
    text,
    color,
    bold: opts.bold || false,
    size: opts.size || 20,
    font: opts.mono ? 'JetBrains Mono' : 'Arial',
    italics: opts.italic || false,
  });
}

function p(children, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after || 0 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    border: opts.borderBottom ? {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: C.type, space: 2 }
    } : undefined,
    tabStops: opts.tabStops || [],
    children: Array.isArray(children) ? children : [children],
  });
}

// Section header: "// ОПЫТ РАБОТЫ" with bottom border
function sectionHeader(title) {
  return p([
    t('// ', C.comment, { size: 22, mono: true }),
    t(title, C.type, { bold: true, size: 22 }),
  ], { before: 280, after: 100, borderBottom: true });
}

// Job block
function job(role, company, period, location) {
  return [
    p([
      t(role, C.black, { bold: true, size: 22 }),
      t('\t', C.comment),
      t(period, C.comment, { size: 18, italic: true }),
    ], {
      before: 160, after: 20,
      tabStops: [{ type: TabStopType.RIGHT, position: 8640 }]
    }),
    p([
      t(company, C.string, { size: 19, bold: true }),
      location ? t('  ·  ' + location, C.comment, { size: 18 }) : t(''),
    ], { after: 60 }),
  ];
}

function bullet(text) {
  return p([
    t('▸  ', C.keyword, { size: 19, mono: true }),
    t(text, C.black, { size: 19 }),
  ], { before: 40, after: 40, indent: 200 });
}

function techLine(text) {
  return p([
    t('// tech: ', C.comment, { size: 17, mono: true }),
    t(text, C.comment, { size: 17, mono: true }),
  ], { before: 30, after: 100, indent: 200 });
}

// Skill row: "Java · Spring Boot · PostgreSQL"
function skillGroup(label, items) {
  return p([
    t(label + ': ', C.purple, { bold: true, size: 19 }),
    t(items.join(' · '), C.black, { size: 19 }),
  ], { before: 50, after: 50 });
}

// Project block
function project(name, url, desc) {
  return [
    p([
      t(name, C.black, { bold: true, size: 20 }),
      url ? t('  github.com/' + url, C.link, { size: 18 }) : t(''),
    ], { before: 120, after: 30 }),
    ...desc.map(d => bullet(d)),
  ];
}

// Contact row item
function contactItem(label, value, color = C.black) {
  return [
    t(label + ': ', C.comment, { size: 18 }),
    t(value, color, { size: 18 }),
    t('   ', C.black, { size: 18 }),
  ];
}

// ── DOCUMENT ─────────────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 20, color: C.black } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // ── HEADER ──
      p([
        t('Фоломеев Алексей', C.black, { bold: true, size: 52 }),
      ], { after: 60 }),

      p([
        t('Fullstack Developer', C.type, { bold: true, size: 26 }),
      ], { after: 80 }),

      // Contacts inline
      p([
        ...contactItem('Email', 'folomeev.aleksei97@gmail.com', C.link),
        ...contactItem('Тел', '+996 556 20 62 03'),
        ...contactItem('GitHub', 'https://github.com/VisionFoland2025', C.link),
        ...contactItem('Город', 'Бишкек'),
      ], { after: 40 }),

      p([
        ...contactItem('LinkedIn', 'https://www.linkedin.com/in/alex-foland/', C.link),
        ...contactItem('Опыт', '2 год'),
        ...contactItem('Английский', 'B2 (Upper-Intermediate)'),
      ], { after: 20 }),

      // Divider paragraph
      p([t('')], {
        before: 60, after: 60,
        borderBottom: true,
      }),

      // ── О СЕБЕ ──
      sectionHeader('О СЕБЕ'),
      p([
        t('Backend-разработчик с 5+ годами опыта создания высоконагруженных систем на Java и Spring. '
        + 'Специализируюсь на микросервисной архитектуре, проектировании REST API и оптимизации баз данных PostgreSQL/Oracle. '
        + 'Опыт внедрения event-driven решений на Kafka, настройки CI/CD и работы с Kubernetes.',
          C.black, { size: 19 })
      ], { before: 80, after: 0 }),

      // ── ОПЫТ РАБОТЫ ──
      sectionHeader('ОПЫТ РАБОТЫ'),

      ...job('Senior Java Developer', 'FinTech Solutions', '2026 — н.в.', 'Бишкек'),
      bullet('Разработка микросервисов на Spring Boot 3.x для обработки платежей — 500 000 транзакций/день'),
      bullet('Проектирование REST API и gRPC-сервисов, документирование через Swagger / OpenAPI'),
      bullet('Оптимизация SQL-запросов PostgreSQL: снижение времени отклика на 40%'),
      bullet('Внедрение Spring Security + JWT + OAuth2, двухфакторная аутентификация'),
      bullet('Event-driven архитектура на Apache Kafka, кэширование через Redis'),
      techLine('Spring Boot 3, PostgreSQL, Kafka, Redis, Docker, Kubernetes, GitHub Actions'),

      ...job('Java Backend Developer', 'RetailTech', '2020 — 2025', 'Бишкек'),
      bullet('Декомпозиция монолита на Spring MVC на микросервисы, покрытие тестами до 85%'),
      bullet('Написание сложных SQL-запросов, хранимых процедур и триггеров в Oracle DB'),
      bullet('Реализация пакетной обработки данных через Spring Batch (5M+ записей/сутки)'),
      bullet('Интеграция с внешними API: 1С, ЭДО, банковские шлюзы через Spring Integration'),
      techLine('Spring MVC, Oracle DB, Spring Batch, RabbitMQ, Hibernate, Jenkins'),

      ...job('Junior Java Developer', 'Softline', '2019 — 2020', 'Бишкек'),
      bullet('Разработка CRUD-модулей на Spring Data JPA и Hibernate'),
      bullet('Написание unit- и интеграционных тестов: JUnit 5, Mockito, Testcontainers'),
      bullet('Работа с MySQL: индексирование таблиц, нормализация схем данных'),
      techLine('Spring MVC, MySQL, JUnit 5, Mockito, Maven, Git'),

      // ── НАВЫКИ ──
      sectionHeader('НАВЫКИ'),
      skillGroup('Backend',    ['Java 17/21', 'Spring Boot', 'Spring Security', 'Spring Data JPA', 'Spring Batch', 'Hibernate']),
      skillGroup('API',        ['REST', 'gRPC', 'WebFlux', 'GraphQL', 'OpenAPI / Swagger']),
      skillGroup('Базы данных',['PostgreSQL', 'MySQL', 'Oracle', 'Redis', 'Flyway', 'Liquibase']),
      skillGroup('Очереди',    ['Apache Kafka', 'RabbitMQ', 'Spring Integration']),
      skillGroup('DevOps',     ['Docker', 'Kubernetes', 'CI/CD', 'Jenkins', 'GitHub Actions', 'SonarQube']),
      skillGroup('Тестирование',['JUnit 5', 'Mockito', 'Testcontainers', 'WireMock']),
      skillGroup('Инструменты',['IntelliJ IDEA', 'Git', 'Maven', 'Gradle', 'Jira', 'Confluence']),

      // ── ПРОЕКТЫ ──
      sectionHeader('ПРОЕКТЫ'),
      ...project(
        'PaymentGateway',
        'apetrov/payment-gateway',
        [
          'Платёжный шлюз с поддержкой 10+ провайдеров, 2M+ транзакций/месяц',
          'Spring Boot + PostgreSQL + Kafka + Redis, deploy в Kubernetes',
        ]
      ),
      ...project(
        'SQLQueryAnalyzer',
        'apetrov/sql-analyzer',
        [
          'Open-source инструмент анализа и оптимизации SQL-запросов — 1200 на GitHub',
          'Парсинг AST запросов, визуализация execution plan, рекомендации по индексам',
        ]
      ),

      // ── ОБРАЗОВАНИЕ ──
      sectionHeader('ОБРАЗОВАНИЕ'),
      p([
        t('ИРНИТУ', C.black, { bold: true, size: 20 }),
        t('\t', C.comment),
        t('2014 — 2019', C.comment, { size: 18, italic: true }),
      ], {
        before: 120, after: 20,
        tabStops: [{ type: TabStopType.RIGHT, position: 8640 }]
      }),

      // ── СЕРТИФИКАТЫ ──
      sectionHeader('СЕРТИФИКАТЫ'),
      bullet('Oracle Certified Professional — Java SE 17 Developer (2023)'),
      bullet('AWS Certified Developer — Associate (2022)'),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('resume_java_ats.docx', buf);
  console.log('Done!');
});
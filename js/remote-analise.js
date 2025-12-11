const DATA_FILES = {
  homolog: 'dados/fornecedores_homologados.xlsx',
  iqf: 'dados/atendimento controle_qualidade.xlsx'
};

const IQF_SEGMENTS = [
  { id: 'excelente', label: 'Excelentes (IQF ≥ 90)', min: 90, max: Infinity },
  { id: 'bom', label: 'Bons (80 ≤ IQF < 90)', min: 80, max: 89.99 },
  { id: 'regular', label: 'Regulares (70 ≤ IQF < 80)', min: 70, max: 79.99 },
  { id: 'critico', label: 'Críticos (IQF < 70)', min: -Infinity, max: 69.99 },
  { id: 'sem-dados', label: 'Sem medição IQF', none: true }
];

const API_KEY_STORAGE_KEY = 'analise-openai-key';
const OPENAI_API_KEY_DEFAULT =
  (typeof process !== 'undefined' && process.env && process.env.OPENAI_API_KEY) ||
  (typeof window !== 'undefined' && window.OPENAI_API_KEY) ||
  '';
const EMAIL_API_ENDPOINT =
  (typeof window !== 'undefined' && window.EMAIL_API_ENDPOINT) ||
  (typeof process !== 'undefined' && process.env && process.env.EMAIL_API_ENDPOINT) ||
  '/api/send-email';
const EMAIL_API_TOKEN =
  (typeof window !== 'undefined' && window.EMAIL_API_TOKEN) ||
  (typeof process !== 'undefined' && process.env && process.env.EMAIL_API_TOKEN) ||
  '';

const ENGEMAN_LOGO_DATA_URI =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAfEAAAH2CAYAAABgN7XFAAAAAXNSR0IArs4c6QAAIABJREFUeF7snQe0JFW1/r9zqqrDjXMnM8QhB8kgIDz1iQ/Fh6gg/MWAIA99IoiBIEmCZBBQFBQVjDwDJswBTJhQJA5hAjAzTM43dXeFc/5rn+qeuTOAcOeG7ur+aq1ZPffe6lP7/Pap+uqkvRV4kAAJkAAJkAAJZJKAyqTVNJoESIAESIAESAAUcTYCEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJJBRAhTxjDqOZpMACZAACZAARZxtgARIgARIgAQySoAinlHH0WwSIAESIAESoIizDZAACZAACZBARglQxDPqOJpNAiRAAiRAAhRxtgESIAESIAESyCgBinhGHUezSYAESIAESIAizjZAAiRAAiRAAhklQBHPqONoNgmQAAmQAAlQxNkGSIAESIAESCCjBCjiGXUczSYBEiABEiABijjbAAmQAAmQAAlklABFPKOOo9kkQAIkQAIkQBFnGyABEiABEiCBjBKgiGfUcTSbBEiABEiABCjibAMkQAIkQAIkkFECFPGMOo5mkwAJkAAJkABFnG2ABEiABEiABDJKgCKeUcfRbBIgARIgARKgiLMNkAAJkAAJkEBGCVDEM+o4mk0CJEACJEACFHG2ARIgARIgARLIKAGKeEYdR7NJgARIgARIgCLONkACJEACJEACGSVAEc+o42g2CZAACZAACVDE2QZIgARIgARIIKMEKOIZdRzNJgESIAESIAGKONsACZAACZAACWSUAEU8o46j2SRAAiRAAiRAEWcbIAESIAESIIGMEqCIZ9RxNJsESIAESIAEKOJsAyRAAiRAAiSQUQIU8Yw6jmaTAAmQAAmQAEWcbYAESIAESIAEMkqAIp5Rx9FsEiABEiABEqCIsw2QAAmQAAmQQEYJUMQz6jiaTQIkQAIkQAIUcbYBEiABEiABEsgoAYp4Rh1Hs0mABEiABEiAIs42QAIkQAIkQAIZJUARz6jjaDYJkAAJkAAJUMTZBkiABEiABEggowQo4hl1HM0mARIgARIgAYo42wAJkAAJkAAJZJQARTyjjqPZJEACJEACJEARZxsgARIgARIggYwSoIhn1HE0mwRIgARIgAQo4mwDJEACJEACJ

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/i;
const MAX_SUPPLIER_SUGGESTIONS = 8;
const MAX_HISTORY_POINTS = 12;
let incrementalId = 0;

const state = {
  suppliers: [],
  supplierById: new Map(),
  segments: [],
  segmentById: new Map(),
  selectedSupplier: null,
  lastFeedback: '',
  lastFeedbackHtml: '',
  lastEmailHtml: '',
  feedbackSupplierId: null,
  monthlySummary: new Map(),
  availableMonths: [],
  emailSubjectTemplate: 'Feedback IQF - {{fornecedor}}',
  openAiApiKey: '',
  inputDefaultPlaceholder: '',
  emailPromptSupplierId: null,
  selectedMonthKey: null,
  monthlySnapshot: null,
  monthlyEmailPromptMonthKey: null,
  monthlySubjectTemplate: 'Relatorio IQF Mensal - {{mes}}',
  lastMonthlyNarrativeHtml: '',
  lastMonthlyEmailHtml: ''
};

const dom = {
  messages: null,
  mainMenu: null,
  userInput: null,
  sendBtn: null,
  toast: null,
  settingsBtn: null,
  settingsOverlay: null,
  settingsCloseBtn: null,
  settingsSaveBtn: null,
  settingsClearBtn: null,
  settingsEmailSubject: null,
  apiKeyBar: null,
  apiKeyInput: null,
  apiKeyToggle: null,
  applyApiKeyBtn: null,
  clearApiKeyBtn: null,
  apiKeyStatus: null,
  emailPromptCard: null,
  emailPromptForm: null,
  emailPromptInput: null,
  emailPromptButton: null,
  emailPromptStatus: null,
  monthlyEmailPromptCard: null,
  monthlyEmailPromptForm: null,
  monthlyEmailPromptInput: null,
  monthlyEmailPromptButton: null,
  monthlyEmailPromptStatus: null
};

function cacheSafePath(path) {
  if (!path) {
    return path;
  }
  const cacheBust = Date.now().toString(36);
  const base =
    typeof window !== 'undefined' && window.location
      ? window.location.origin && window.location.origin !== 'null'
        ? window.location.origin
        : window.location.href
      : '';
  try {
    const resolved = new URL(path, base || undefined);
    resolved.searchParams.set('cb', cacheBust);
    return resolved.toString();
  } catch (error) {
    const [cleanPath, ...rest] = String(path).split('?');
    const encodedPath = encodeURI(cleanPath);
    const query = rest.length ? '?' + rest.join('?') : '';
    const separator = query ? '&' : '?';
    return encodedPath + query + separator + 'cb=' + cacheBust;
  }
}

function getOpenAiApiKey() {
  return (state.openAiApiKey || OPENAI_API_KEY_DEFAULT || '').trim();
}

function setOpenAiApiKey(value, options) {
  const normalized = value ? String(value).trim() : '';
  const previous = state.openAiApiKey;
  state.openAiApiKey = normalized;
  try {
    if (normalized) {
      localStorage.setItem(API_KEY_STORAGE_KEY, normalized);
    } else {
      localStorage.removeItem(API_KEY_STORAGE_KEY);
    }
  } catch (error) {
    console.warn('[analise:setOpenAiApiKey]', error);
  }
  if (dom.apiKeyInput && dom.apiKeyInput.value !== normalized) {
    dom.apiKeyInput.value = normalized;
  }
  updateApiKeyStatus();
  const changed = previous !== normalized;
  if (options?.showToast && changed) {
    showToast(normalized ? 'Chave OpenAI salva.' : 'Chave OpenAI removida.');
  }
  return changed;
}

function updateApiKeyStatus() {
  if (!dom.apiKeyStatus) {
    return;
  }
  const key = getOpenAiApiKey();
  if (key) {
    const masked =
      key.length > 12 ? key.slice(0, 4) + '...' + key.slice(-4) : 'chave ativa';
    dom.apiKeyStatus.textContent = 'Chave ativa (' + masked + ')';
    dom.apiKeyStatus.dataset.status = 'ready';
  } else {
    dom.apiKeyStatus.textContent = 'Nenhuma chave informada';
    dom.apiKeyStatus.dataset.status = 'empty';
  }
}

function toggleApiKeyVisibility() {
  if (!dom.apiKeyInput || !dom.apiKeyToggle) {
    return;
  }
  const hidden = dom.apiKeyInput.type === 'password';
  dom.apiKeyInput.type = hidden ? 'text' : 'password';
  dom.apiKeyToggle.classList.toggle('active', hidden);
  dom.apiKeyToggle.setAttribute('aria-label', hidden ? 'Ocultar chave' : 'Mostrar chave');
}

function handleApplyApiKey() {
  if (!dom.apiKeyInput) {
    return;
  }
  const changed = setOpenAiApiKey(dom.apiKeyInput.value, { showToast: true });
  if (!changed) {
    updateApiKeyStatus();
    showToast(getOpenAiApiKey() ? 'Chave OpenAI ativa.' : 'Nenhuma chave informada.');
  }
}

function handleClearApiKey() {
  const changed = setOpenAiApiKey('', { showToast: true });
  if (!changed) {
    updateApiKeyStatus();
  }
  if (dom.apiKeyInput) {
    dom.apiKeyInput.focus();
  }
}

function handleApiKeyBlur() {
  if (!dom.apiKeyInput) {
    return;
  }
  setOpenAiApiKey(dom.apiKeyInput.value, { showToast: false });
}

function handleApiKeyKeydown(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    handleApplyApiKey();
  } else if (event.key === 'Escape') {
    event.preventDefault();
    if (dom.apiKeyInput) {
      dom.apiKeyInput.value = state.openAiApiKey;
    }
  }
}

document.addEventListener('DOMContentLoaded', init);

window.showSubmenu = handleShowSubmenu;
window.showIndicadoresMensais = handleIndicadoresMensais;
window.contactBase = handleContactBase;
window.showProcedimentoEngeman = handleProcedimento;
window.previewEmailTemplate = previewEmailTemplate;
window.openSupplierSearch = handleSupplierSearch;

async function init() {
  cacheDom();
  bindEvents();
  loadSettings();

  const loadingMessage = appendBotMessage('Carregando planilhas de fornecedores, aguarde...');
  try {
    await loadData();
    updateMessage(
      loadingMessage,
      'Bases carregadas com sucesso! Acesse "Indicadores dos Fornecedores" para explorar os agrupamentos por IQF.',
      true
    );
    showToast('Dados atualizados. Escolha um agrupamento para iniciar a analise.');
  } catch (error) {
    console.error('[analise:init]', error);
    updateMessage(
      loadingMessage,
      'Nao foi possivel carregar as planilhas. Verifique os arquivos em "dados/" e recarregue a página.',
      true
    );
  }
}

function cacheDom() {
  dom.messages = document.getElementById('messagesContainer');
  dom.mainMenu = document.getElementById('main-menu');
  dom.userInput = document.getElementById('userInput');
  dom.sendBtn = document.getElementById('sendBtn');
  if (dom.userInput && !state.inputDefaultPlaceholder) {
    state.inputDefaultPlaceholder = dom.userInput.placeholder || '';
  }
  dom.toast = document.getElementById('toast');
  dom.settingsBtn = document.getElementById('settingsBtn');
  dom.settingsOverlay = document.getElementById('settingsOverlay');
  dom.settingsCloseBtn = document.getElementById('closeSettingsBtn');
  dom.settingsSaveBtn = document.getElementById('saveSettingsBtn');
  dom.settingsClearBtn = document.getElementById('clearSettingsBtn');
  dom.settingsEmailSubject = document.getElementById('emailSubjectInput');
  dom.apiKeyBar = document.getElementById('apiKeyBar');
  dom.apiKeyInput = document.getElementById('apiKeyInput');
  dom.apiKeyToggle = document.getElementById('toggleApiKeyBtn');
  dom.applyApiKeyBtn = document.getElementById('applyApiKeyBtn');
  dom.clearApiKeyBtn = document.getElementById('clearApiKeyBtn');
  dom.apiKeyStatus = document.getElementById('apiKeyStatus');
  updateApiKeyStatus();
}

function bindEvents() {
  if (dom.sendBtn) {
    dom.sendBtn.addEventListener('click', handleUserSend);
  }
  if (dom.userInput) {
    dom.userInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        handleUserSend();
      }
    });
  }
  if (dom.settingsBtn) {
    dom.settingsBtn.addEventListener('click', openSettings);
  }
  if (dom.settingsCloseBtn) {
    dom.settingsCloseBtn.addEventListener('click', closeSettings);
  }
  if (dom.settingsOverlay) {
    dom.settingsOverlay.addEventListener('click', (event) => {
      if (event.target === dom.settingsOverlay) {
        closeSettings();
      }
    });
  }
  if (dom.settingsSaveBtn) {
    dom.settingsSaveBtn.addEventListener('click', saveSettings);
  }
  if (dom.settingsClearBtn) {
    dom.settingsClearBtn.addEventListener('click', clearSettings);
  }
  if (dom.apiKeyInput) {
    dom.apiKeyInput.addEventListener('keydown', handleApiKeyKeydown);
    dom.apiKeyInput.addEventListener('blur', handleApiKeyBlur);
  }
  if (dom.applyApiKeyBtn) {
    dom.applyApiKeyBtn.addEventListener('click', handleApplyApiKey);
  }
  if (dom.clearApiKeyBtn) {
    dom.clearApiKeyBtn.addEventListener('click', handleClearApiKey);
  }
  if (dom.apiKeyToggle) {
    dom.apiKeyToggle.addEventListener('click', toggleApiKeyVisibility);
  }

}

function loadSettings() {
  let persistedApiKey = '';
  try {
    const persistedSubject = localStorage.getItem('analise-email-subject');
    if (persistedSubject) {
      state.emailSubjectTemplate = persistedSubject;
      if (dom.settingsEmailSubject) {
        dom.settingsEmailSubject.value = persistedSubject;
      }
    } else if (dom.settingsEmailSubject) {
      dom.settingsEmailSubject.value = state.emailSubjectTemplate;
    }
    persistedApiKey = localStorage.getItem(API_KEY_STORAGE_KEY) || '';
  } catch (error) {
    console.warn('[analise:loadSettings]', error);
  }
  if (persistedApiKey) {
    state.openAiApiKey = persistedApiKey;
  } else if (OPENAI_API_KEY_DEFAULT) {
    state.openAiApiKey = OPENAI_API_KEY_DEFAULT;
  }
  if (dom.apiKeyInput) {
    dom.apiKeyInput.value = state.openAiApiKey;
    dom.apiKeyInput.type = 'password';
  }
  updateApiKeyStatus();
}

async function loadData() {
  if (typeof XLSX === 'undefined') {
    throw new Error('Biblioteca XLSX nao carregada');
  }
  const [homRows, iqfRows] = await Promise.all([
    loadWorkbook(DATA_FILES.homolog),
    loadWorkbook(DATA_FILES.iqf)
  ]);
  const homolog = homRows.map(mapHomolog).filter(Boolean);
  const iqf = iqfRows
    .map(mapIqf)
    .filter((row) => row.code || row.name || row.iqf !== null || row.occ);
  buildSupplierState(homolog, iqf);
}

async function loadWorkbook(path) {
  const response = await fetch(cacheSafePath(path), { cache: 'no-store' });
  if (!response.ok) {
    throw new Error('Falha ao carregar ' + path + ' (' + response.status + ')');
  }
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
  const rows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    if (worksheet) {
      rows.push(...XLSX.utils.sheet_to_json(worksheet, { defval: null }));
    }
  });
  return rows;
}

function buildSupplierState(homologRows, iqfRows) {
  const codeBuckets = new Map();
  const nameBuckets = new Map();

  iqfRows.forEach((row) => {
    if (row.code) {
      const key = String(row.code);
      if (!codeBuckets.has(key)) {
        codeBuckets.set(key, []);
      }
      codeBuckets.get(key).push(row);
    }
    if (row.name) {
      const normalized = normalizeText(row.name);
      if (normalized) {
        if (!nameBuckets.has(normalized)) {
          nameBuckets.set(normalized, []);
        }
        nameBuckets.get(normalized).push(row);
      }
    }
  });

  const suppliers = [];
  const usedIds = new Set();
  const seenKeys = new Set();

  const registerSupplier = (supplier) => {
    if (!supplier || usedIds.has(supplier.id)) {
      return;
    }
    usedIds.add(supplier.id);
    suppliers.push(supplier);
  };

  homologRows.forEach((hom) => {
    const codeKey = hom.code ? String(hom.code) : null;
    const nameKey = hom.name ? normalizeText(hom.name) : null;
    const byCode = codeKey && codeBuckets.get(codeKey) ? codeBuckets.get(codeKey) : [];
    const byName = nameKey && nameBuckets.get(nameKey) ? nameBuckets.get(nameKey) : [];
    const rows = byCode.length >= byName.length ? byCode : byName;
    const supplier = composeSupplier(hom, rows, usedIds);
    registerSupplier(supplier);
    if (codeKey) {
      seenKeys.add('code:' + codeKey);
    }
    if (nameKey) {
      seenKeys.add('name:' + nameKey);
    }
  });

  codeBuckets.forEach((rows, code) => {
    if (!seenKeys.has('code:' + code)) {
      const placeholderHom = { code, name: rows[0]?.name || null, status: 'Pendente', score: null, expire: null };
      registerSupplier(composeSupplier(placeholderHom, rows, usedIds));
      seenKeys.add('code:' + code);
    }
  });

  nameBuckets.forEach((rows, normalizedName) => {
    if (!seenKeys.has('name:' + normalizedName)) {
      const placeholderHom = { code: null, name: rows[0]?.name || null, status: 'Pendente', score: null, expire: null };
      registerSupplier(composeSupplier(placeholderHom, rows, usedIds));
      seenKeys.add('name:' + normalizedName);
    }
  });

  suppliers.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'accent' }));
  state.suppliers = suppliers;
  state.supplierById = new Map(suppliers.map((supplier) => [supplier.id, supplier]));
  state.segments = prepareSegments(suppliers);
  state.segmentById = new Map(state.segments.map((segment) => [segment.id, segment]));
  buildMonthlySummaries(iqfRows, suppliers);
}

function composeSupplier(hom, iqfRows, usedIds) {
  const iqfSummary = summarizeIqf(iqfRows || []);
  const code = hom?.code ? String(hom.code) : iqfRows?.[0]?.code ? String(iqfRows[0].code) : null;
  const name = safeSupplierName(hom?.name, iqfRows?.[0]?.name);
  const normalizedName = normalizeText(name);
  const homologScore = roundValue(hom?.score);
  const baseStatus = hom?.status || 'Pendente';
  const status = deriveStatus(baseStatus, iqfSummary.average, homologScore);
  const id = buildSupplierId(code, normalizedName, usedIds);

  return {
    id,
    code,
    name,
    normalizedName,
    status,
    baseStatus,
    homologScore,
    averageIqf: iqfSummary.average,
    iqfSamples: iqfSummary.samples,
    lastIqfDate: iqfSummary.lastDate,
    expire: hom?.expire || null,
    document: iqfSummary.document,
    occurrences: iqfSummary.occurrences,
    iqfHistory: iqfSummary.history,
    searchText: [name || '', code || ''].join(' ').toLowerCase(),
    baseStatusRaw: hom?.status || null
  };
}

function summarizeIqf(rows) {
  const history = [];
  const occurrences = [];
  let document = '';

  rows.forEach((row) => {
    if (row.iqf !== null) {
      history.push({
        value: roundValue(row.iqf),
        date: row.date,
        formattedDate: formatDate(row.date)
      });
    }
    if (row.occ) {
      occurrences.push({
        text: row.occ,
        date: row.date,
        formattedDate: formatDate(row.date),
        document: row.document,
        severity: severityLevel(row.occ)
      });
    }
    if (!document && row.document) {
      document = row.document;
    }
  });

  history.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  occurrences.sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  const trimmedHistory = history.slice(0, MAX_HISTORY_POINTS);
  const samples = history.length;
  const average = samples ? roundValue(history.reduce((sum, entry) => sum + (entry.value || 0), 0) / samples) : null;
  const lastDate = history.length ? history[0].date : null;

  return { history: trimmedHistory, occurrences, average, samples, lastDate, document };
}

function safeSupplierName(...names) {
  const options = names.filter(Boolean).map((value) => safeString(value));
  const preferred = options.find((value) => value.length > 0);
  return preferred || 'Fornecedor nao identificado';
}

function prepareSegments(suppliers) {
  return IQF_SEGMENTS.map((segment) => {
    const matches = suppliers.filter((supplier) => {
      if (segment.none) {
        return supplier.averageIqf === null;
      }
      if (supplier.averageIqf === null) {
        return false;
      }
      return supplier.averageIqf >= segment.min && supplier.averageIqf <= segment.max;
    });
    return { ...segment, suppliers: matches };
  }).filter((segment) => segment.suppliers.length);
}

function buildMonthlySummaries(iqfRows, suppliers) {
  const supplierIndex = new Map();
  suppliers.forEach((supplier) => {
    if (supplier.code) {
      supplierIndex.set('code:' + String(supplier.code).trim(), supplier);
    }
    if (supplier.normalizedName) {
      supplierIndex.set('name:' + supplier.normalizedName, supplier);
    }
  });

  const summary = new Map();

  iqfRows.forEach((row) => {
    if (!row || row.iqf === null || row.iqf === undefined || !row.date) {
      return;
    }
    const monthKey = row.date.slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(monthKey)) {
      return;
    }
    if (!summary.has(monthKey)) {
      summary.set(monthKey, { totalSum: 0, totalCount: 0, suppliers: new Map() });
    }
    const monthEntry = summary.get(monthKey);
    monthEntry.totalSum += row.iqf;
    monthEntry.totalCount += 1;

    const normalizedName = normalizeText(row.name);
    let supplierKey = null;
    if (row.code) {
      supplierKey = 'code:' + String(row.code).trim();
    } else if (normalizedName) {
      supplierKey = 'name:' + normalizedName;
    } else {
      supplierKey = 'name:' + safeString(row.name).toLowerCase();
    }

    if (!monthEntry.suppliers.has(supplierKey)) {
      const supplierRef =
        supplierIndex.get(supplierKey) ||
        (row.code ? supplierIndex.get('code:' + String(row.code).trim()) : null) ||
        (normalizedName ? supplierIndex.get('name:' + normalizedName) : null) ||
        null;
      monthEntry.suppliers.set(supplierKey, {
        key: supplierKey,
        name: safeSupplierName(row.name),
        code: supplierRef?.code || row.code || null,
        status: supplierRef?.status || 'Pendente',
        sum: 0,
        count: 0
      });
    }
    const supplierEntry = monthEntry.suppliers.get(supplierKey);
    supplierEntry.sum += row.iqf;
    supplierEntry.count += 1;
  });

  state.monthlySummary = summary;
  state.availableMonths = Array.from(summary.keys()).sort((a, b) => b.localeCompare(a));
}

function handleShowSubmenu(key) {
  if (key !== 'fornecedores') {
    appendBotMessage('Menu em construção. Em breve novas opções aqui.', true);
    return;
  }
  if (!state.suppliers.length) {
    appendBotMessage('Os dados ainda estão sendo carregados. Aguarde alguns instantes e tente novamente.', true);
    return;
  }
  renderIqfSegments();
}

function handleSupplierSearch() {
  if (!dom.userInput) {
    appendBotMessage('Campo de busca indisponivel no momento.', true);
    return;
  }
  const content = createMessage('bot');
  content.innerHTML =
    '<p><strong>Pesquisar fornecedor</strong></p>' +
    '<p>Digite o nome completo ou parte do nome no campo inferior e pressione Enter. O sistema listará as correspondências mais próximas.</p>';
  dom.userInput.placeholder = 'Pesquisar fornecedor por nome...';
  dom.userInput.value = '';
  dom.userInput.focus();
  showToast('Digite o nome do fornecedor e pressione Enter para pesquisar.');
}

function renderMainMenuOptions() {
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = '<strong>Selecione uma op&ccedil;&atilde;o para continuar:</strong>';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  const options = [
    { label: 'Pesquisar Fornecedor', handler: handleSupplierSearch },
    { label: 'Indicadores dos Fornecedores', handler: () => handleShowSubmenu('fornecedores') },
    { label: 'Indicadores Mensais', handler: handleIndicadoresMensais },
    { label: 'Procedimento Engeman', handler: handleProcedimento },
    { label: 'Contato Base', handler: handleContactBase },
    { label: 'Configura&ccedil;&otilde;es', handler: openSettings }
  ];
  options.forEach((option) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = option.label;
    button.addEventListener('click', () => option.handler());
    actions.appendChild(button);
  });
  content.appendChild(actions);
}

function renderIqfSegments() {
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = 'Escolha um agrupamento de fornecedores pela média IQF:';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  state.segments.forEach((segment) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.innerHTML = segment.label + ' (' + segment.suppliers.length + ')';
    button.addEventListener('click', () => showSuppliersBySegment(segment.id));
    actions.appendChild(button);
  });
  content.appendChild(actions);
}

function showSuppliersBySegment(segmentId) {
  const segment = state.segmentById.get(segmentId);
  if (!segment) {
    appendBotMessage('Segmento nao localizado. Atualize a página e tente novamente.', true);
    return;
  }
  const content = createMessage('bot');
  const title = document.createElement('p');
  title.innerHTML = '<strong>' + segment.label + '</strong> — ' + segment.suppliers.length + ' fornecedor(es). Selecione para detalhar:';
  content.appendChild(title);

  if (!segment.suppliers.length) {
    const empty = document.createElement('p');
    empty.textContent = 'Nenhum fornecedor neste agrupamento.';
    content.appendChild(empty);
    return;
  }

  const list = document.createElement('div');
  list.className = 'message-actions supplier-grid';
  const sortedSuppliers = segment.suppliers
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'accent' }));
  sortedSuppliers.forEach((supplier) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = supplier.name + (supplier.code ? ' (' + supplier.code + ')' : '');
    button.addEventListener('click', () => showSupplierDetails(supplier.id));
    list.appendChild(button);
  });
  content.appendChild(list);
  state.currentSegmentId = segmentId;
}

function showSupplierDetails(supplierId) {
  const supplier = state.supplierById.get(supplierId);
  if (!supplier) {
    appendBotMessage('Fornecedor nao encontrado. Utilize o menu novamente.', true);
    return;
  }
  state.selectedSupplier = supplier;
  state.feedbackSupplierId = null;
  state.lastFeedback = '';
  state.lastFeedbackHtml = '';
  state.lastEmailHtml = '';
  clearEmailPrompt();

  const content = createMessage('bot');
  const header = document.createElement('div');
  header.innerHTML =
    '<h3>' +
    escapeHtml(supplier.name) +
    '</h3>' +
    renderStatusBadge(supplier.status) +
    '<p>Código: <strong>' +
    (supplier.code ? escapeHtml(supplier.code) : 'n/d') +
    '</strong> • Nota Homologação: <strong>' +
    (supplier.homologScore !== null ? supplier.homologScore : 'n/d') +
    '</strong> • Média IQF: <strong>' +
    (supplier.averageIqf !== null ? supplier.averageIqf : 'n/d') +
    '</strong></p>' +
    '<p>Registros IQF: <strong>' +
    supplier.iqfSamples +
    '</strong> • Última medição: <strong>' +
    (formatDate(supplier.lastIqfDate) || 'n/d') +
    '</strong></p>';
  content.appendChild(header);

  if (supplier.expire) {
    const expire = document.createElement('p');
    expire.textContent = 'Validade da homologação: ' + formatDate(supplier.expire);
    content.appendChild(expire);
  }

  const feedbackId = safeDomId('feedback', supplier.id);
  const feedbackCard = document.createElement('div');
  feedbackCard.className = 'feedback-card';
  feedbackCard.id = feedbackId;
  content.appendChild(feedbackCard);

  const refreshLink = document.createElement('button');
  refreshLink.className = 'inline-link';
  refreshLink.type = 'button';
  refreshLink.textContent = 'Reprocessar analise';
  refreshLink.addEventListener('click', () => {
    generateSupplierFeedback(supplier, feedbackCard);
  });
  feedbackCard.appendChild(refreshLink);
  feedbackCard.__refreshButton = refreshLink;

  const occurrencesCard = document.createElement('div');
  occurrencesCard.className = 'occurrences-card';
  occurrencesCard.innerHTML = '<h4>Ocorrências relevantes</h4>';
  if (supplier.occurrences.length) {
    const list = document.createElement('ul');
    supplier.occurrences.forEach((occ) => {
      const item = document.createElement('li');
      item.innerHTML =
        '<strong>' +
        (occ.formattedDate || 'Data n/d') +
        ':</strong> ' +
        escapeHtml(occ.text) +
        (occ.document ? ' (Doc: ' + escapeHtml(occ.document) + ')' : '');
      list.appendChild(item);
    });
    occurrencesCard.appendChild(list);
  } else {
    const empty = document.createElement('p');
    empty.textContent = 'Sem ocorrências registradas para este fornecedor.';
    occurrencesCard.appendChild(empty);
  }
  content.appendChild(occurrencesCard);

  if (supplier.iqfHistory.length) {
    const historyCard = document.createElement('div');
    historyCard.className = 'occurrences-card';
    historyCard.innerHTML = '<h4>Histórico de IQF</h4>';
    const historyList = document.createElement('ul');
    supplier.iqfHistory.forEach((entry) => {
      const item = document.createElement('li');
      item.textContent = (entry.formattedDate || 'Data n/d') + ' \u2013 IQF: ' + entry.value;
      historyList.appendChild(item);
    });
    historyCard.appendChild(historyList);
    content.appendChild(historyCard);
  }

  const hint = document.createElement('p');
  hint.className = 'message-hint';
  hint.innerHTML =
    'Utilize o cartao "Envio automatico do feedback" para informar o e-mail de contato e disparar a analise.';
  content.appendChild(hint);

  generateSupplierFeedback(supplier, feedbackCard);
  renderEmailPrompt(supplier);
}

function buildSupplierClassificationNarrative(supplier) {
  const iqf = Number.isFinite(supplier.averageIqf) ? supplier.averageIqf : null;
  const formattedIqf = iqf !== null ? formatScoreValue(iqf) : 'N/D';
  const samples = supplier.iqfSamples || 0;
  const lastMeasurement = supplier.lastIqfDate ? formatDate(supplier.lastIqfDate) : 'n/d';
  const expireDate = supplier.expire ? formatDate(supplier.expire) : null;
  const criteriaAtencao = [
    'Cumprimento de prazos conforme o pedido de compra ou contrato.',
    'Comunicação, garantia e suporte pós-venda.',
    'Qualidade do material ou serviço entregue.',
    'Conformidade com os itens descritos no pedido ou contrato.'
  ];
  const criteriaCriticos = [
    'Cumprimento de prazos conforme o pedido de compra ou contrato.',
    'Conformidade com os itens descritos no pedido ou contrato.',
    'Qualidade do material ou serviço entregue.',
    'Comunicação, garantia e suporte pós-venda.',
    'Embalagem e identificação do material.',
    'Cumprimento das normas de segurança.',
    'Envio de documentos obrigatórios (boleto, notas fiscais e certificados necessarios).'
  ];

  let message = '';
  if (iqf === null) {
    message =
      'Ainda nao há medições registradas para este fornecedor. Solicitar atualização do ciclo de avaliação.';
  } else if (iqf > 75) {
    message =
      'Agradecemos pela parceria e informamos que a empresa obteve excelente performance na avaliação mais recente.' +
      '\n\n📊 Nota IQF: ' +
      formattedIqf +
      '\n🏆 Classificação: Aprovado - desempenho de excelência.' +
      '\n\nEsse resultado reforça o comprometimento com qualidade, prazos e conformidade. Mantemos confiança na continuidade desta parceria sólida.';
  } else if (iqf >= 70) {
    const temas = pickRandom(criteriaAtencao, 2);
    message =
      'Compartilhamos o resultado da avaliação periódica de desempenho.' +
      '\n\n📊 Nota IQF: ' +
      formattedIqf +
      '\n⚠️ Classificação: Em atenção - desempenho no limite mínimo.' +
      '\n\nRecomendamos foco nos seguintes aspectos para evitar regressões:\n' +
      temas.map((tema) => '- ' + tema).join('\n') +
      '\n\nA manutenção dos indicadores é essencial para preservar a parceria.';
  } else {
    const temas = pickRandom(criteriaCriticos, 3);
    message =
      'Informamos que o fornecedor foi reprovado no Índice de Qualidade (IQF) e precisa de ações imediatas.' +
      '\n\n📊 Nota IQF: ' +
      formattedIqf +
      '\n❌ Classificação: Reprovado - abaixo do padrão mínimo (70,00).' +
      '\n\nFalhas observadas:\n' +
      temas.map((tema) => '- ' + tema).join('\n') +
      '\n\nSolicitamos analise interna das nao conformidades e plano corretivo. A reincidência pode impactar futuros fornecimentos.';
  }

  const occSummary = summarizeOccurrenceTexts(supplier.occurrences);
  if (occSummary) {
    message += '\n\n🔴 Ocorrências registradas no atendimento:\n' + occSummary;
  }

  message +=
    '\n\n📑 Indicadores complementares:\n' +
    '- Avaliações consideradas: ' +
    samples +
    '\n' +
    '- Ultima avaliação IQF: ' +
    lastMeasurement;
  if (expireDate) {
    message += '\n- Validade da homologacao: ' + expireDate;
  }

  message +=
    '\n\nLegenda de notas:\n' +
    '- 0 a 69: Reprovado - nao atingiu os critérios mínimos de qualidade e conformidade.\n' +
    '- A partir de 70: Aprovado - desempenho dentro dos parâmetros estabelecidos.';

  return message;
}

function summarizeOccurrenceTexts(occurrences) {
  if (!occurrences || !occurrences.length) {
    return '';
  }
  const unique = [];
  const seen = new Set();
  occurrences.forEach((occ) => {
    const text = safeString((occ && (occ.text || occ.occ)) || '');
    if (text && !seen.has(text)) {
      seen.add(text);
      unique.push(text);
    }
  });
  if (!unique.length) {
    return '';
  }
  const lines = unique.slice(0, 5).map((entry) => '• ' + entry);
  if (unique.length > lines.length) {
    lines.push('• ... outras ocorrências foram registradas no periodo.');
  }
  return lines.join('\n');
}

function generateSupplierFeedback(supplier, container) {
  if (!container) {
    return;
  }
  const baselineText = buildSupplierClassificationNarrative(supplier);
  const baselineHtml = formatFeedback(baselineText);
  renderFeedbackCard(container, {
    title: 'Analise executiva',
    subtitle: 'Resumo gerado com dados do IQF',
    icon: '??',
    bodyHtml: baselineHtml,
    hint: 'Gerando insights adicionais com IA...'
  });
  appendRefreshControl(container);
  state.lastFeedback = baselineText;
  state.lastFeedbackHtml = baselineHtml;
  state.lastEmailHtml = '';
  state.feedbackSupplierId = supplier.id;
  const apiKey = getOpenAiApiKey();
  if (!apiKey) {
    renderFeedbackCard(container, {
      title: 'Analise executiva',
      subtitle: 'Resumo gerado com dados do IQF',
      icon: '??',
      bodyHtml: baselineHtml,
      hint: 'Informe a chave da API OpenAI no painel de configuracoes (icone de engrenagem) para complementar a analise automaticamente.'
    });
    appendRefreshControl(container);
    showToast('Abra o painel de configuracoes (icone de engrenagem) e cole a chave da API OpenAI para habilitar os insights automaticos.');
    if (dom.apiKeyInput) {
      dom.apiKeyInput.focus();
    }
    return;
  }

  fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + apiKey
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: 'Você é um especialista em qualificação de fornecedores.' },
        { role: 'user', content: buildSupplierPrompt(supplier) }
      ]
    })
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao consultar a API');
        });
      }
      return response.json();
    })
    .then((payload) => {
      if (state.feedbackSupplierId !== supplier.id) {
        return;
      }
      const answer = payload?.choices?.[0]?.message?.content?.trim();
      if (!answer) {
        throw new Error('Resposta vazia da API');
      }
      const combined = baselineText + '\n\n' + answer;
      const formatted = formatFeedback(combined);
      state.lastFeedback = combined;
      state.lastFeedbackHtml = formatted;
      state.lastEmailHtml = '';
      state.feedbackSupplierId = supplier.id;
      renderFeedbackCard(container, {
        title: 'Analise executiva',
        subtitle: 'Complementada com IA (GPT-4o mini)',
        icon: '🤖',
        bodyHtml: formatted
      });
      appendRefreshControl(container);
    })
    .catch((error) => {
      console.error('[analise:gpt]', error);
      state.lastFeedback = baselineText;
      state.lastFeedbackHtml = baselineHtml;
      state.lastEmailHtml = '';
      renderFeedbackCard(container, {
        title: 'Analise executiva',
        subtitle: 'Resumo gerado com dados do IQF',
        icon: '⚠️',
        bodyHtml: baselineHtml,
        hint: 'Nao foi possivel complementar a analise com IA agora. Tente novamente em instantes.'
      });
      appendRefreshControl(container);
      showToast('Erro ao consultar a API OpenAI. Tente novamente em instantes.');
    });
}

function buildSupplierPrompt(supplier) {
  const iqfLines = supplier.iqfHistory.length
    ? supplier.iqfHistory
        .map((entry) => '- ' + (entry.formattedDate || 'Data n/d') + ': IQF ' + entry.value)
        .join('\n')
    : '- sem histórico registrado';
  const occLines = supplier.occurrences.length
    ? supplier.occurrences
        .slice(0, 6)
        .map((occ) => '- ' + (occ.formattedDate || 'Data n/d') + ': ' + occ.text)
        .join('\n')
    : '- sem ocorrências registradas';

  const statusAlert =
    'Alerta de risco: ' +
    (supplier.status === 'Reprovado'
      ? 'Fornecedor reprovado - detalhe claramente todas as falhas e acione plano corretivo imediato.'
      : supplier.status === 'Pendente' || supplier.status === 'Em atencao'
      ? 'Fornecedor em atencao - destaque impactos potenciais e orientacoes preventivas.'
      : 'Fornecedor homologado - manter tom reconhecendo desempenho, mas reforce pontos de melhoria.');

  const situationalGuidance = [
    supplier.status === 'Reprovado'
      ? 'Em "Riscos e Pontos de Atencao", explicite as consequencias de reprovação e indique reforço de governança.'
      : supplier.status === 'Pendente' || supplier.status === 'Em atencao'
      ? 'Em "Riscos e Pontos de Atencao", sinalize urgência moderada e medidas preventivas para evitar escalonamento.'
      : 'Em "Pontos Fortes", reconheça o desempenho e proponha melhorias incrementais para manter a maturidade.'
  ].join(' ');

  return [
    'Voce atua como consultor senior de supply chain da Engeman. Produza uma analise executiva em portugues do Brasil com base apenas nos dados abaixo.',
    statusAlert,
    '',
    'Perfil do fornecedor:',
    '- Nome: ' + supplier.name,
    '- Codigo interno: ' + (supplier.code || 'n/d'),
    '- Status consolidado: ' + supplier.status + ' (status de origem: ' + (supplier.baseStatusRaw || 'n/d') + ')',
    '- Nota de homologacao vigente: ' + (supplier.homologScore !== null ? supplier.homologScore : 'n/d'),
    '- Media IQF consolidada: ' + (supplier.averageIqf !== null ? supplier.averageIqf : 'sem avaliacao'),
    '- Total de avaliacoes IQF: ' + supplier.iqfSamples,
    '- Ultima avaliacao IQF registrada: ' + (formatDate(supplier.lastIqfDate) || 'n/d'),
    '- Validade da homologacao: ' + (formatDate(supplier.expire) || 'n/d'),
    '',
    'Historico recente de IQF (ordem cronologica decrescente):',
    iqfLines,
    '',
    'Ocorrencias relevantes registradas:',
    occLines,
    '',
    'Estruture a resposta seguindo exatamente o formato abaixo, usando frases curtas com metricas numericas sempre que possivel:',
    'Visao Geral:',
    '- ...',
    'Pontos Fortes:',
    '- ...',
    'Riscos e Pontos de Atencao:',
    '- ...',
    'Recomendacoes Taticas:',
    '- ...',
    'Plano de Monitoramento (30 dias):',
    '- ...',
    'Observacoes Complementares:',
    '- ...',
    situationalGuidance,
    '',
    'Nao invente informacoes. Quando algum dado nao estiver disponivel, registre essa ausencia.'
  ].join('\n');
}

function applyInlineFormatting(text) {
  return escapeHtml(text).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
}

const HEADING_EMOJI_MAP = {
  'visao geral': '📊',
  'pontos de atencao': '⚠️',
  'pontos fortes': '•',
  'riscos e pontos de atencao': '•',
  'recomendacoes taticas': '🛠️',
  'plano de monitoramento (30 dias)': '•',
  'observacoes complementares': '•',
  'fornecedores reprovados': '•',
  'acoes imediatas': '•',
  'conclusao': '•',
  'alertas prioritarios': '❗'
};

function decorateHeadingLabel(label) {
  if (!label) {
    return label;
  }
  const base = label.trim();
  let normalized = base.toLowerCase();
  if (typeof base.normalize === 'function') {
    normalized = base
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase();
  }
  const emoji = HEADING_EMOJI_MAP[normalized];
  if (!emoji || base.startsWith(emoji)) {
    return base;
  }
  return emoji + ' ' + base;
}

function formatFeedback(text) {
  const lines = String(text || '').split(/\r?\n/);
  const html = [];
  let listBuffer = [];

  const flushList = () => {
    if (listBuffer.length) {
      html.push('<ul>' + listBuffer.join('') + '</ul>');
      listBuffer = [];
    }
  };

  lines.forEach((raw) => {
    const line = raw.trim();
    if (!line) {
      flushList();
      return;
    }
    if (/^(?:[-*\u2022\u0007])\s+/.test(line)) {
      const item = applyInlineFormatting(line.replace(/^(?:[-*\u2022\u0007])\s*/, '').trim());
      listBuffer.push('<li>' + item + '</li>');
      return;
    }
    flushList();
    if (/^[A-Za-z0-9].*:\s*$/.test(line)) {
      const rawHeading = line.replace(/:$/, '').trim();
      const heading = applyInlineFormatting(decorateHeadingLabel(rawHeading));
      html.push('<h4>' + heading + '</h4>');
      return;
    }
    html.push('<p>' + applyInlineFormatting(line) + '</p>');
  });
  flushList();
  return html.join('');
}
function renderFeedbackCard(container, config) {
  if (!container) {
    return;
  }
  const title = config?.title ? escapeHtml(config.title) : 'Analise executiva';
  const subtitle = config?.subtitle ? escapeHtml(config.subtitle) : '';
  const icon = escapeHtml(config?.icon || '🤖');
  const bodyHtml = config?.bodyHtml || '<p>Nenhum conteudo disponível.</p>';
  const hintText = config?.hint ? escapeHtml(config.hint) : null;
  container.innerHTML =
    '<div class="ai-card-header">' +
    '<span class="ai-chip" aria-hidden="true">' +
    icon +
    '</span>' +
    '<div class="ai-card-titles">' +
    '<p class="ai-card-title">' +
    title +
    '</p>' +
    (subtitle ? '<p class="ai-card-subtitle">' + subtitle + '</p>' : '') +
    '</div>' +
    '</div>' +
    '<div class="ai-card-body">' +
    bodyHtml +
    '</div>' +
    (hintText ? '<p class="ai-hint">' + hintText + '</p>' : '');
}

function appendRefreshControl(container) {
  if (!container || !container.__refreshButton) {
    return;
  }
  container.appendChild(container.__refreshButton);
}

function buildEmailNarrative(supplier) {
  const iqf = Number.isFinite(supplier.averageIqf) ? supplier.averageIqf : null;
  const parts = [];

  if (iqf === null) {
    parts.push(
      '<p>Ainda nao identificamos medições recentes de IQF para este fornecedor. Recomendamos atualizar a avaliação para manter o controle de desempenho.</p>'
    );
  } else if (iqf > 75) {
    parts.push(
      '<p>Agradecemos pela parceria e informamos que sua empresa obteve excelente performance na nossa avaliação mais recente.</p>',
      '<p>📊 Nota IQF: <strong>' + formatScoreValue(iqf) + '</strong><br>🏆 Classificação: Aprovado - Desempenho de excelência.</p>',
      '<p>Esse resultado demonstra alto nível de comprometimento com qualidade, prazos e conformidade. Seguimos confiantes na continuidade desta parceria sólida e eficiente.</p>',
      '<p>Qualquer dúvida, estamos à disposição.</p>'
    );
  } else if (iqf >= 70) {
    const criteriosFalhos = randomSample(
      [
        'Cumprimento de prazos conforme o pedido de compra ou contrato.',
        'Comunicação, garantia e suporte pós-venda.',
        'Qualidade do material ou serviço entregue.',
        'Conformidade com os itens descritos no pedido ou contrato.'
      ],
      2
    );
    parts.push(
      '<p>Compartilhamos abaixo o resultado da avaliação periódica de desempenho:</p>',
      '<p>📊 Nota IQF: <strong>' + formatScoreValue(iqf) + '</strong><br>⚠️ Nota mínima exigida: 70,00</p>',
      '<p>Embora tecnicamente aprovado, o resultado indica performance no limite mínimo aceitável. Recomendamos atenção especial aos seguintes aspectos:</p>',
      '<ul>' + criteriosFalhos.map((c) => '<li>' + escapeHtml(c) + '</li>').join('') + '</ul>',
      '<p>A manutenção de bons indicadores é essencial para seguirmos com uma parceria de confiança e excelência.</p>'
    );
  } else {
    const criteriosFalhos = randomSample(
      [
        'Cumprimento de prazos conforme o pedido de compra ou contrato.',
        'Conformidade com os itens descritos no pedido ou contrato.',
        'Qualidade do material ou serviço entregue.',
        'Comunicação, garantia e suporte pós-venda.',
        'Embalagem e identificação do material.',
        'Cumprimento das normas de segurança.',
        'Envio de documentos obrigatórios (boleto, NF-e, certificados).'
      ],
      3
    );
    parts.push(
      '<p>Informamos que, conforme a avaliação periódica de desempenho, sua empresa foi reprovada no Índice de Qualidade do Fornecedor (IQF):</p>',
      '<p>📊 Nota IQF: <strong>' + formatScoreValue(iqf) + '</strong><br>❌ Classificação: Reprovado - Abaixo do padrão mínimo (70,00)</p>',
      '<p>A reprovação ocorreu devido a falhas nos seguintes critérios:</p>',
      '<ul>' + criteriosFalhos.map((c) => '<li>' + escapeHtml(c) + '</li>').join('') + '</ul>',
      '<p>Solicitamos analise interna das nao conformidades e implementação de medidas corretivas. A reincidência pode impactar futuros fornecimentos.</p>'
    );
  }

  const occurrencesHtml = formatOccurrencesHtml(supplier.occurrences);
  if (occurrencesHtml) {
    parts.push(occurrencesHtml);
  }

  parts.push(
    '<h4>Legendas de notas</h4>',
    '<ul>' +
      '<li><strong>0 a 69: Reprovado</strong> - indica que o fornecedor nao atingiu os critérios mínimos de qualidade e conformidade.</li>' +
      '<li><strong>A partir de 70: Aprovado</strong> - indica desempenho satisfatório conforme os critérios estabelecidos.</li>' +
      '</ul>',
    '<p>Em caso de apontamentos negativos, pedimos a analise e correção. A reincidência de problemas pode suspender o fornecedor da Engeman.</p>',
    '<p>Seguimos confiantes na continuidade desta parceria sólida e eficiente.</p>',
    '<p>Atenciosamente,<br>Equipe de Suprimentos</p>'
  );

  return { html: parts.join('') };
}

function randomSample(list, count) {
  const pool = list.slice();
  const picks = [];
  const target = Math.min(count, pool.length);
  while (picks.length < target && pool.length) {
    const index = Math.floor(Math.random() * pool.length);
    picks.push(pool.splice(index, 1)[0]);
  }
  return picks;
}

function formatOccurrencesHtml(occurrences) {
  if (!occurrences || !occurrences.length) {
    return '';
  }
  const items = occurrences.slice(0, 5).map((occ) => {
    const dateLabel = occ.formattedDate || formatDate(occ.date) || 'Data nao informada';
    const text = escapeHtml(occ.text || occ.occ || '');
    return '<li><strong>' + escapeHtml(dateLabel) + ':</strong> ' + text + '</li>';
  });
  return '<p>🔴 <strong>Ocorrências registradas no atendimento:</strong></p><ul>' + items.join('') + '</ul>';
}

function buildEmailHtml(supplier, analysisHtml) {
  const supplierDisplayName = supplier?.name && supplier.name.trim() ? supplier.name : null;
  const supplierName = escapeHtml(supplierDisplayName || (supplier.code ? 'Fornecedor ' + supplier.code : 'Fornecedor nao identificado'));
  const statusLabel = escapeHtml(supplier?.status || 'Pendente');
  const statusClass = getStatusBadgeClass(supplier?.status);
  const lastIqfDate = formatDate(supplier?.lastIqfDate) || 'N/D';
  const iqfScore = Number.isFinite(supplier?.averageIqf) ? formatScoreValue(supplier.averageIqf) : 'N/D';
  const lastEvaluation = iqfScore !== 'N/D' ? lastIqfDate + ' · IQF ' + iqfScore : lastIqfDate;
  const feedbackSection = analysisHtml || '<p class="empty">Feedback nao disponível no momento.</p>';
  const observationDetails = [
    { label: 'Média IQF atual', value: iqfScore !== 'N/D' ? iqfScore : null },
    {
      label: 'Avaliações consideradas',
      value: supplier?.iqfSamples ? supplier.iqfSamples + ' registros analisados' : null
    },
    { label: 'Status consolidado', value: supplier?.status ? supplier.status : null },
    { label: 'Última avaliação IQF', value: lastEvaluation !== 'N/D' ? lastEvaluation : null },
    { label: 'Validade da homologação', value: supplier?.expire ? formatDate(supplier.expire) : null }
  ].filter((item) => item.value);
  const observationsSection = observationDetails.length
    ? '<ul class="observation-list">' +
      observationDetails
        .map((item) => '<li><strong>' + escapeHtml(item.label) + ':</strong> ' + escapeHtml(item.value) + '</li>')
        .join('') +
      '</ul>'
    : '<p class="empty">Sem observações complementares no momento.</p>';
  const occurrenceItems = (supplier?.occurrences || []).slice(0, 5).map((occ) => {
    const dateLabel = occ.formattedDate || formatDate(occ.date) || 'Data nao informada';
    const text = escapeHtml(occ.text || occ.occ || 'Sem descricao.');
    const document = occ.document ? ' (Doc: ' + escapeHtml(occ.document) + ')' : '';
    return '<li><strong>' + escapeHtml(dateLabel) + '</strong><span>' + text + document + '</span></li>';
  });
  const occurrencesSection = occurrenceItems.length ? '<ul class="occurrence-list">' + occurrenceItems.join('') + '</ul>' : '';

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Atualizacao Engeman - ${supplierName}</title>
<style>
body{margin:0;background:#f4f5f7;font-family:"Segoe UI",Arial,sans-serif;color:#0f172a;}
.wrapper{padding:30px 16px;}
.content{max-width:640px;margin:0 auto;background:#ffffff;border-radius:18px;box-shadow:0 20px 40px rgba(15,23,42,0.08);overflow:hidden;}
.header{padding:32px 30px;background:#0f172a;color:#f8fafc;text-align:center;}
.header-logo{height:44px;margin:0 auto 14px;display:block;}
.header h1{margin:0;font-size:22px;}
.header p{margin:6px 0 0;font-size:13px;color:rgba(248,250,252,0.8);}
.card{padding:26px 30px;border-top:1px solid #e2e8f0;}
.card:first-of-type{border-top:none;}
.card h2{margin:0 0 14px;font-size:17px;}
.info-list{list-style:none;padding:0;margin:0;}
.info-list li{margin-bottom:12px;}
.info-list span{display:block;font-size:12px;text-transform:uppercase;letter-spacing:0.05em;color:#94a3b8;}
.info-list strong{display:block;font-size:16px;margin-top:3px;color:#0f172a;}
.status-pill{display:inline-block;margin-top:4px;padding:6px 14px;border-radius:999px;font-size:12px;font-weight:600;color:#fff;}
.badge-homologado{background:#10b981;}
.badge-reprovado{background:#dc2626;}
.badge-pendente{background:#d97706;}
.badge-em-atencao{background:#f97316;}
.feedback-card h2{margin-bottom:8px;}
.feedback-columns{display:flex;flex-wrap:wrap;gap:18px;margin-top:12px;}
.feedback-column{flex:1 1 260px;background:#f8fafc;border-radius:16px;padding:18px;}
.feedback-column h3{margin:0 0 10px;font-size:15px;}
.feedback-column ul{list-style:none;padding:0;margin:0;}
.feedback-column li+li{margin-top:8px;}
.feedback-column strong{color:#0f172a;}
.occurrence-list{list-style:none;padding:0;margin:0;}
.occurrence-list li{padding:10px 0;border-bottom:1px solid #f1f5f9;}
.occurrence-list li:last-child{border-bottom:none;}
.occurrence-list strong{display:block;font-size:13px;color:#475569;}
.occurrence-list span{display:block;font-size:14px;color:#0f172a;margin-top:2px;}
.empty{font-size:13px;color:#94a3b8;}
.footer{padding:18px 30px;background:#f8fafc;font-size:12px;color:#64748b;text-align:center;}
</style>
</head>
<body>
  <div class="wrapper">
    <div class="content">
      <div class="header">
        <img src="${ENGEMAN_LOGO_DATA_URI}" alt="Logo Engeman" class="header-logo">
        <h1>Atualizacao de desempenho do fornecedor</h1>
        <p>Mensagem automatica da Central de Suprimentos Engeman</p>
      </div>
      <div class="card">
        <h2>Dados principais</h2>
        <ul class="info-list">
          <li>
            <span>Fornecedor</span>
            <strong>${supplierName}</strong>
          </li>
          <li>
            <span>Status consolidado</span>
            <span class="status-pill ${statusClass}">${statusLabel}</span>
          </li>
          <li>
            <span>Ultima avaliacao registrada</span>
            <strong>${lastEvaluation}</strong>
          </li>
        </ul>
      </div>
      <div class="card feedback-card">
        <h2>Feedback detalhado do fornecedor</h2>
        <div class="feedback-columns">
          <div class="feedback-column feedback-content">
            ${feedbackSection}
          </div>
          <div class="feedback-column feedback-observations">
            <h3>Observações complementares</h3>
            ${observationsSection}
          </div>
        </div>
      </div>
      ${occurrencesSection ? `<div class="card"><h2>Ocorrências recentes</h2>${occurrencesSection}</div>` : ''}
      <div class="footer">
        Este e-mail foi enviado automaticamente. Em caso de duvidas, contate o time de Suprimentos.
      </div>
    </div>
  </div>
</body>
</html>`;
}

function buildMonthlyEmailHtml(snapshot, narrativeHtml) {
  if (!snapshot) {
    return '';
  }
  const monthLabelRaw = snapshot.monthLabel || formatMonthLabel(snapshot.monthKey);
  const monthLabel = escapeHtml(monthLabelRaw);
  const totals = snapshot.totals || {};
  const distribution = snapshot.distribution || {};
  const averageLabel = Number.isFinite(totals.globalAverage) ? formatScoreValue(totals.globalAverage) : 'N/D';
  const totalSamples = totals.totalSamples ?? 0;
  const totalSuppliers = totals.totalSuppliers ?? 0;
  const narrative = narrativeHtml || '<p class="empty">Resumo mensal indisponivel.</p>';
  const generatedDateIso = snapshot.generatedAt ? String(snapshot.generatedAt).slice(0, 10) : new Date().toISOString().slice(0, 10);
  const generatedLabel = formatDate(generatedDateIso) || generatedDateIso;

  const buildItemsList = (items, emptyMessage) => {
    if (!items || !items.length) {
      return '<p class="empty">' + escapeHtml(emptyMessage) + '</p>';
    }
    const rows = items.map((item) => {
      const name = escapeHtml(item.name || 'Fornecedor n/d');
      const status = escapeHtml(item.status || 'Pendente');
      const avg = Number.isFinite(item.avg) ? formatScoreValue(item.avg) : 'N/D';
      const count = item.count ?? 0;
      return '<li><strong>' + name + ' (' + status + ')</strong><span>IQF ' + avg + ' • ' + count + ' avaliacoes</span></li>';
    });
    return '<ul class="occurrence-list">' + rows.join('') + '</ul>';
  };

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatorio Mensal - ${monthLabel}</title>
<style>
body{margin:0;background:#f4f5f7;font-family:"Segoe UI",Arial,sans-serif;color:#0f172a;}
.wrapper{padding:30px 16px;}
.content{max-width:640px;margin:0 auto;background:#ffffff;border-radius:18px;box-shadow:0 20px 40px rgba(15,23,42,0.08);overflow:hidden;}
.header{padding:32px 30px;background:#111827;color:#f8fafc;text-align:center;}
.header-logo{height:44px;margin:0 auto 14px;display:block;}
.header h1{margin:0;font-size:22px;}
.header p{margin:6px 0 0;font-size:13px;color:rgba(248,250,252,0.75);}
.card{padding:24px 30px;border-top:1px solid #e2e8f0;}
.card:first-of-type{border-top:none;}
.card h2{margin:0 0 14px;font-size:17px;}
.stats-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;}
.stat{background:#f8fafc;border-radius:14px;padding:14px;}
.stat span{display:block;font-size:12px;text-transform:uppercase;letter-spacing:0.05em;color:#94a3b8;}
.stat strong{display:block;font-size:18px;margin-top:4px;}
.distribution{display:flex;gap:16px;flex-wrap:wrap;}
.distribution div{flex:1 1 140px;background:#f1f5f9;border-radius:12px;padding:12px;text-align:center;}
.distribution span{display:block;font-size:12px;color:#475569;}
.distribution strong{display:block;font-size:20px;margin-top:4px;}
.occurrence-list{list-style:none;padding:0;margin:0;}
.occurrence-list li{padding:10px 0;border-bottom:1px solid #f1f5f9;}
.occurrence-list li:last-child{border-bottom:none;}
.occurrence-list strong{display:block;font-size:13px;color:#475569;}
.occurrence-list span{display:block;font-size:14px;color:#0f172a;margin-top:2px;}
.empty{font-size:13px;color:#94a3b8;}
.footer{padding:18px 30px;background:#f8fafc;font-size:12px;color:#64748b;text-align:center;}
</style>
</head>
<body>
  <div class="wrapper">
    <div class="content">
      <div class="header">
        <img src="${ENGEMAN_LOGO_DATA_URI}" alt="Logo Engeman" class="header-logo">
        <h1>Relatorio IQF Mensal</h1>
        <p>Periodo avaliado: ${monthLabel}</p>
      </div>
      <div class="card">
        <h2>Resumo executivo</h2>
        <div class="stats-grid">
          <div class="stat">
            <span>Média global IQF</span>
            <strong>${averageLabel}</strong>
          </div>
          <div class="stat">
            <span>Avaliações analisadas</span>
            <strong>${totalSamples}</strong>
          </div>
          <div class="stat">
            <span>Fornecedores avaliados</span>
            <strong>${totalSuppliers}</strong>
          </div>
          <div class="stat">
            <span>Consolidado gerado em</span>
            <strong>${escapeHtml(generatedLabel || 'N/D')}</strong>
          </div>
        </div>
        <div class="distribution">
          <div>
            <span>Aprovados</span>
            <strong>${distribution.aprovados ?? 0}</strong>
          </div>
          <div>
            <span>Em atenção</span>
            <strong>${distribution.emAtencao ?? 0}</strong>
          </div>
          <div>
            <span>Reprovados</span>
            <strong>${distribution.reprovados ?? 0}</strong>
          </div>
        </div>
      </div>
      <div class="card">
        <h2>Resumo da analise</h2>
        ${narrative}
      </div>
      <div class="card">
        <h2>Fornecedores reprovados</h2>
        ${buildItemsList(snapshot.reprovados, 'Nenhum fornecedor reprovado neste periodo.')}
      </div>
      <div class="card">
        <h2>Fornecedores em atenção</h2>
        ${buildItemsList(snapshot.emAtencao, 'Nenhum fornecedor em atenção neste período.')}
      </div>
      <div class="card">
        <h2>Destaques de excelência</h2>
        ${buildItemsList(snapshot.excelencia, 'Nenhum destaque de excelência registrado para este mês.')}
      </div>
      <div class="card">
        <h2>Panorama resumido</h2>
        ${buildItemsList(snapshot.panorama, 'Sem dados suficientes para compor o panorama.')}
      </div>
      <div class="footer">
        Este relatório foi enviado automaticamente pela Central de Suprimentos Engeman.
      </div>
    </div>
  </div>
</body>
</html>`;
}

function getStatusBadgeClass(status) {
  if (status === 'Homologado') {
    return 'badge-homologado';
  }
  if (status === 'Reprovado') {
    return 'badge-reprovado';
  }
  return 'badge-pendente';
}

function formatScoreValue(value) {
  if (!Number.isFinite(value)) {
    return 'N/D';
  }
  return value.toFixed(2);
}

function pickRandom(list, count) {
  const source = Array.isArray(list) ? list.slice() : [];
  const picks = [];
  const limit = Math.min(count, source.length);
  for (let i = 0; i < limit; i += 1) {
    const index = Math.floor(Math.random() * source.length);
    picks.push(source.splice(index, 1)[0]);
  }
  return picks;
}

function htmlToPlainText(html) {
  const cleaned = String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(?:p|div|section|tr|h[1-6])>/gi, '\n')
    .replace(/<li>/gi, '- ')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/?[^>]+>/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return decodeBasicEntities(cleaned);
}

function decodeBasicEntities(text) {
  const entities = {
    ' ': ' ',
    '&': '&',
    '<': '<',
    '>': '>',
    '"': '"',
    '&#39;': "'",
    'á': 'a',
    'à': 'a',
    'ã': 'a',
    'ç': 'c',
    'é': 'e',
    'ê': 'e',
    'í': 'i',
    'ó': 'o',
    'ô': 'o',
    'õ': 'o',
    'ú': 'u',
    '&uuml;': 'u'
  };
  return String(text || '').replace(/&[a-z#0-9]+;/gi, (entity) => entities[entity.toLowerCase()] || entity);
}


function tryCopyHtml(html) {
  if (!navigator.clipboard || typeof navigator.clipboard.writeText !== 'function') {
    return Promise.resolve(false);
  }
  return navigator.clipboard.writeText(html).then(
    () => true,
    () => false
  );
}

function previewEmailTemplate() {
  if (!state.lastEmailHtml) {
    showToast('Gere o feedback e informe o e-mail para criar o layout completo.');
    return;
  }
  const encoded = 'data:text/html;charset=utf-8,' + encodeURIComponent(state.lastEmailHtml);
  window.open(encoded, '_blank', 'noopener,noreferrer');
}

function handleUserSend() {
  if (!dom.userInput) {
    return;
  }
  const message = dom.userInput.value.trim();
  if (!message) {
    dom.userInput.placeholder = state.inputDefaultPlaceholder || dom.userInput.placeholder;
    return;
  }
  dom.userInput.value = '';
  dom.userInput.placeholder = state.inputDefaultPlaceholder || dom.userInput.placeholder;
  appendUserMessage(message);

  if (EMAIL_REGEX.test(message)) {
    sendSupplierEmail(message);
    return;
  }

  handleSupplierQuery(message);
}

function handleSupplierQuery(query) {
  const normalized = normalizeText(query);
  if (!normalized) {
    appendBotMessage('Informe o nome ou código do fornecedor que deseja visualizar.', true);
    return;
  }
  if (['menu', 'opcoes', 'opcao', 'inicio', 'voltar', 'principal'].includes(normalized)) {
    renderMainMenuOptions();
    showToast('Menu principal exibido novamente.');
    return;
  }
  const exact = state.suppliers.find(
    (supplier) => supplier.code === query || supplier.normalizedName === normalized
  );
  if (exact) {
    showSupplierDetails(exact.id);
    return;
  }
  const matches = state.suppliers
    .filter((supplier) => supplier.searchText.includes(query.toLowerCase()))
    .slice(0, MAX_SUPPLIER_SUGGESTIONS);
  if (!matches.length) {
    appendBotMessage(
      'Nenhum fornecedor encontrado para "' + escapeHtml(query) + '". Utilize o menu de agrupamentos por IQF para navegar.',
      true
    );
    return;
  }
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = 'Fornecedores relacionados a <strong>' + escapeHtml(query) + '</strong>:';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  matches.forEach((supplier) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = supplier.name + (supplier.code ? ' (' + supplier.code + ')' : '');
    button.addEventListener('click', () => showSupplierDetails(supplier.id));
    actions.appendChild(button);
  });
  content.appendChild(actions);
}

function clearEmailPrompt() {
  if (dom.emailPromptCard?.isConnected) {
    const wrapper = dom.emailPromptCard.closest('.message');
    if (wrapper?.parentNode) {
      wrapper.parentNode.removeChild(wrapper);
    } else {
      dom.emailPromptCard.remove();
    }
  } else if (dom.messages) {
    const fallbackCard = dom.messages.querySelector('.email-capture-card');
    fallbackCard?.closest('.message')?.remove();
  }
  dom.emailPromptCard = null;
  dom.emailPromptForm = null;
  dom.emailPromptInput = null;
  dom.emailPromptButton = null;
  dom.emailPromptStatus = null;
  state.emailPromptSupplierId = null;
}

function renderEmailPrompt(supplier) {
  if (!supplier || !dom.messages) {
    return;
  }
  if (state.emailPromptSupplierId === supplier.id && dom.emailPromptCard?.isConnected) {
    return;
  }

  clearEmailPrompt();

  const content = createMessage('bot');
  const card = document.createElement('div');
  card.className = 'email-capture-card';

  const title = document.createElement('h4');
  title.textContent = 'Envio automático do feedback';
  card.appendChild(title);

  const displayName =
    supplier.name && supplier.name.trim()
      ? supplier.name
      : supplier.code
      ? 'Fornecedor ' + supplier.code
      : 'o fornecedor selecionado';

  const description = document.createElement('p');
  description.innerHTML =
    'Informe o e-mail de contato de <strong>' +
    escapeHtml(displayName) +
    '</strong> para disparar a analise diretamente do painel.';
  card.appendChild(description);

  const form = document.createElement('form');
  form.className = 'email-capture-form';

  const input = document.createElement('input');
  input.type = 'email';
  input.name = 'supplier-email';
  input.placeholder = 'contato@fornecedor.com.br';
  input.autocomplete = 'email';
  input.inputMode = 'email';
  input.required = true;
  input.maxLength = 120;

  const submit = document.createElement('button');
  submit.type = 'submit';
  submit.className = 'email-capture-submit';
  submit.textContent = 'Enviar e-mail';

  const status = document.createElement('span');
  status.className = 'email-capture-status';

  const canSend = Boolean(EMAIL_API_ENDPOINT);
  if (!canSend) {
    input.disabled = true;
    submit.disabled = true;
    status.textContent = 'Configure o servidor de envio para habilitar o disparo automático.';
    status.dataset.state = 'error';
  } else {
    status.textContent = 'Digite o e-mail e enviaremos automaticamente o feedback.';
    status.dataset.state = 'idle';
  }

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const value = input.value.trim();
    input.setCustomValidity('');
    if (!value) {
      input.focus();
      return;
    }
    if (!EMAIL_REGEX.test(value)) {
      input.setCustomValidity('Informe um e-mail válido.');
      input.reportValidity();
      return;
    }
    sendSupplierEmail(value);
  });

  input.addEventListener('input', () => {
    input.setCustomValidity('');
    if (status.dataset.state === 'error' && EMAIL_API_ENDPOINT) {
      status.textContent = 'Digite o e-mail e enviaremos automaticamente o feedback.';
      status.dataset.state = 'idle';
    }
  });

  form.appendChild(input);
  form.appendChild(submit);
  card.appendChild(form);
  card.appendChild(status);
  content.appendChild(card);

  dom.emailPromptCard = card;
  dom.emailPromptForm = form;
  dom.emailPromptInput = input;
  dom.emailPromptButton = submit;
  dom.emailPromptStatus = status;
  state.emailPromptSupplierId = supplier.id;

  if (canSend) {
    setTimeout(() => {
      dom.emailPromptInput?.focus();
    }, 150);
  }
}

function setEmailPromptStatus(message, status, supplierId) {
  if (!dom.emailPromptStatus) {
    return;
  }
  if (supplierId && state.emailPromptSupplierId !== supplierId) {
    return;
  }
  dom.emailPromptStatus.textContent = message;
  dom.emailPromptStatus.dataset.state = status || 'idle';
}

function handleEmailPromptSuccess(email, supplierId) {
  if (state.emailPromptSupplierId !== supplierId) {
    return;
  }
  if (dom.emailPromptInput) {
    dom.emailPromptInput.value = '';
  }
  setEmailPromptStatus('E-mail enviado para ' + email + '.', 'success', supplierId);
  setTimeout(() => {
    if (state.emailPromptSupplierId === supplierId && dom.emailPromptStatus?.dataset.state === 'success') {
      dom.emailPromptStatus.textContent = 'Envie novamente caso precise reenviar o feedback.';
      dom.emailPromptStatus.dataset.state = 'idle';
    }
  }, 5000);
}

function handleEmailPromptError(supplierId) {
  if (state.emailPromptSupplierId !== supplierId) {
    return;
  }
  setEmailPromptStatus(
    'Nao foi possivel enviar. Verifique o servidor de e-mail e tente novamente.',
    'error',
    supplierId
  );
}

function clearMonthlyEmailPrompt() {
  if (dom.monthlyEmailPromptCard?.isConnected) {
    dom.monthlyEmailPromptCard.remove();
  }
  dom.monthlyEmailPromptCard = null;
  dom.monthlyEmailPromptForm = null;
  dom.monthlyEmailPromptInput = null;
  dom.monthlyEmailPromptButton = null;
  dom.monthlyEmailPromptStatus = null;
  state.monthlyEmailPromptMonthKey = null;
}

function renderMonthlyEmailPrompt(container, monthKey, monthLabel) {
  if (!container) {
    return;
  }
  clearMonthlyEmailPrompt();

  const card = document.createElement('div');
  card.className = 'email-capture-card';

  const title = document.createElement('h4');
  title.textContent = 'Envio automático do relatório mensal';
  card.appendChild(title);

  const description = document.createElement('p');
  const labelText = monthLabel || formatMonthLabel(monthKey);
  description.innerHTML =
    'Informe o e-mail do gestor para enviar automaticamente o consolidado de <strong>' +
    escapeHtml(labelText) +
    '</strong>.';
  card.appendChild(description);

  const form = document.createElement('form');
  form.className = 'email-capture-form';

  const input = document.createElement('input');
  input.type = 'email';
  input.name = 'monthly-email';
  input.placeholder = 'gestor@empresa.com';
  input.autocomplete = 'email';
  input.required = true;
  input.maxLength = 120;

  const submit = document.createElement('button');
  submit.type = 'submit';
  submit.className = 'email-capture-submit';
  submit.textContent = 'Enviar relatório';

  const status = document.createElement('span');
  status.className = 'email-capture-status';

  form.appendChild(input);
  form.appendChild(submit);
  card.appendChild(form);
  card.appendChild(status);
  container.appendChild(card);

  dom.monthlyEmailPromptCard = card;
  dom.monthlyEmailPromptForm = form;
  dom.monthlyEmailPromptInput = input;
  dom.monthlyEmailPromptButton = submit;
  dom.monthlyEmailPromptStatus = status;
  state.monthlyEmailPromptMonthKey = monthKey;

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const value = input.value.trim();
    input.setCustomValidity('');
    if (!value) {
      input.focus();
      return;
    }
    if (!EMAIL_REGEX.test(value)) {
      input.setCustomValidity('Informe um e-mail válido.');
      input.reportValidity();
      return;
    }
    sendMonthlyEmail(value);
  });

  input.addEventListener('input', () => {
    input.setCustomValidity('');
  });

  refreshMonthlyEmailPromptState();
}

function refreshMonthlyEmailPromptState() {
  if (!dom.monthlyEmailPromptStatus || !dom.monthlyEmailPromptInput || !dom.monthlyEmailPromptButton) {
    return;
  }
  const monthMatches = state.monthlyEmailPromptMonthKey && state.selectedMonthKey === state.monthlyEmailPromptMonthKey;
  const canSendEndpoint = Boolean(EMAIL_API_ENDPOINT);
  const hasNarrative = Boolean(state.lastMonthlyNarrativeHtml && monthMatches);

  if (!canSendEndpoint) {
    dom.monthlyEmailPromptInput.disabled = true;
    dom.monthlyEmailPromptButton.disabled = true;
    dom.monthlyEmailPromptStatus.textContent =
      'Configure o endpoint de e-mail para liberar o envio do relatório mensal.';
    dom.monthlyEmailPromptStatus.dataset.state = 'error';
    return;
  }

  if (!monthMatches) {
    dom.monthlyEmailPromptInput.disabled = true;
    dom.monthlyEmailPromptButton.disabled = true;
    dom.monthlyEmailPromptStatus.textContent =
      'Selecione um mês para habilitar o envio do relatório correspondente.';
    dom.monthlyEmailPromptStatus.dataset.state = 'idle';
    return;
  }

  if (!hasNarrative) {
    dom.monthlyEmailPromptInput.disabled = true;
    dom.monthlyEmailPromptButton.disabled = true;
    dom.monthlyEmailPromptStatus.textContent =
      'Aguarde o resumo com IA finalizar para liberar o envio automático.';
    dom.monthlyEmailPromptStatus.dataset.state = 'idle';
    return;
  }

  dom.monthlyEmailPromptInput.disabled = false;
  dom.monthlyEmailPromptButton.disabled = false;
  dom.monthlyEmailPromptStatus.textContent = 'Digite o e-mail e enviaremos o relatório mensal automaticamente.';
  dom.monthlyEmailPromptStatus.dataset.state = 'idle';
}

function setMonthlyEmailStatus(message, status) {
  if (!dom.monthlyEmailPromptStatus) {
    return;
  }
  dom.monthlyEmailPromptStatus.textContent = message;
  dom.monthlyEmailPromptStatus.dataset.state = status || 'idle';
}

function handleMonthlyEmailSuccess(email) {
  if (!dom.monthlyEmailPromptInput || !dom.monthlyEmailPromptStatus) {
    return;
  }
  dom.monthlyEmailPromptInput.value = '';
  setMonthlyEmailStatus('Relatório enviado para ' + email + '.', 'success');
  setTimeout(() => {
    if (dom.monthlyEmailPromptStatus?.dataset.state === 'success') {
      setMonthlyEmailStatus('Envie novamente caso precise reenviar o relatório.', 'idle');
    }
  }, 5000);
}

function handleMonthlyEmailError() {
  setMonthlyEmailStatus('Nao foi possivel enviar o relatório mensal. Tente novamente mais tarde.', 'error');
}

function sendSupplierEmail(email) {
  if (!state.selectedSupplier) {
    appendBotMessage('Selecione um fornecedor antes de informar o e-mail para envio.', true);
    return;
  }
  if (!state.lastFeedbackHtml || state.feedbackSupplierId !== state.selectedSupplier.id) {
    appendBotMessage(
      'Gere o feedback com a IA antes de enviar o e-mail. Clique em "Reprocessar analise" caso precise atualizar o conteudo.',
      true
    );
    return;
  }
  if (!email || !EMAIL_REGEX.test(email)) {
    appendBotMessage('Informe um endereco de e-mail valido para prosseguir com o envio automatico.', true);
    return;
  }
  if (!EMAIL_API_ENDPOINT) {
    appendBotMessage(
      'O envio automatico de e-mail nao esta configurado. Verifique se o servidor de envio esta em execucao.',
      true
    );
    setEmailPromptStatus(
      'Envio automatico indisponivel. Configure o servidor antes de prosseguir.',
      'error',
      state.selectedSupplier.id
    );
    return;
  }

  const supplier = state.selectedSupplier;
  const subjectTemplate = state.emailSubjectTemplate || 'Feedback IQF - {{fornecedor}}';
  const displayName = supplier.name && supplier.name.trim() ? supplier.name : supplier.code ? 'Fornecedor ' + supplier.code : 'Fornecedor';
  let subject = subjectTemplate.includes('{{fornecedor}}')
    ? subjectTemplate.replace(/{{fornecedor}}/gi, displayName)
    : subjectTemplate + ' - ' + displayName;

  const analysisHtml = state.lastFeedbackHtml || '<p>Analise da IA indisponivel.</p>';
  const emailHtml = buildEmailHtml(supplier, analysisHtml);
  const plainTextBody = htmlToPlainText(emailHtml);
  state.lastEmailHtml = emailHtml;
  setEmailPromptStatus('Enviando e-mail automaticamente...', 'sending', supplier.id);

  const payload = {
    to: email,
    subject,
    html: emailHtml,
    text: plainTextBody,
    supplier: {
      id: supplier.id,
      name: supplier.name,
      code: supplier.code,
      status: supplier.status,
      averageIqf: supplier.averageIqf,
      homologScore: supplier.homologScore
    },
    generatedAt: new Date().toISOString()
  };

  showToast('Enviando e-mail automatico...');
  tryCopyHtml(emailHtml).catch(() => {});

  const headers = { 'Content-Type': 'application/json' };
  if (EMAIL_API_TOKEN) {
    headers.Authorization = 'Bearer ' + EMAIL_API_TOKEN;
  }

  fetch(EMAIL_API_ENDPOINT, {
    method: 'POST',
    headers,
    body: JSON.stringify(payload)
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao enviar o e-mail.');
        });
      }
      return response.json().catch(() => ({}));
    })
    .then(() => {
      showToast('E-mail enviado com sucesso!');
      appendBotMessage(
        'E-mail enviado automaticamente para <strong>' +
          escapeHtml(email) +
          '</strong>. <button type="button" class="inline-link" onclick="previewEmailTemplate()">Ver layout completo</button>',
        true
      );
      handleEmailPromptSuccess(email, supplier.id);
    })
    .catch((error) => {
      console.error('[analise:send-email]', error);
      showToast('Falha ao enviar o e-mail. Verifique as configuracoes.');
      appendBotMessage(
        'Nao foi possivel concluir o envio automatico. Verifique a configuracao do endpoint de e-mail e tente novamente.',
        true
      );
      handleEmailPromptError(supplier.id);
    });
}

function sendMonthlyEmail(email) {
  if (!state.selectedMonthKey || !state.monthlySnapshot) {
    appendBotMessage('Selecione um mês nos indicadores mensais para preparar o relatório antes do envio.', true);
    return;
  }
  if (!state.lastMonthlyNarrativeHtml) {
    appendBotMessage(
      'Aguarde o resumo com IA concluir para enviar o relatório mensal automaticamente.',
      true
    );
    return;
  }
  if (!EMAIL_API_ENDPOINT) {
    appendBotMessage(
      'O envio automático de e-mail não está configurado. Verifique se o servidor está em execução.',
      true
    );
    handleMonthlyEmailError();
    return;
  }

  const snapshot = state.monthlySnapshot;
  const monthLabel = snapshot.monthLabel || formatMonthLabel(snapshot.monthKey);
  const template = state.monthlySubjectTemplate || 'Relatorio IQF Mensal - {{mes}}';
  const subject = template.includes('{{mes}}')
    ? template.replace(/{{mes}}/gi, monthLabel)
    : template + ' - ' + monthLabel;

  const emailHtml = buildMonthlyEmailHtml(snapshot, state.lastMonthlyNarrativeHtml);
  const plainTextBody = htmlToPlainText(emailHtml);

  state.lastMonthlyEmailHtml = emailHtml;
  state.lastEmailHtml = emailHtml;
  setMonthlyEmailStatus('Enviando relatório mensal...', 'sending');

  const payload = {
    to: email,
    subject,
    html: emailHtml,
    text: plainTextBody,
    reportType: 'monthly',
    month: snapshot.monthKey,
    monthLabel: snapshot.monthLabel,
    totals: snapshot.totals,
    distribution: snapshot.distribution,
    generatedAt: new Date().toISOString()
  };

  const headers = { 'Content-Type': 'application/json' };
  if (EMAIL_API_TOKEN) {
    headers.Authorization = 'Bearer ' + EMAIL_API_TOKEN;
  }

  fetch(EMAIL_API_ENDPOINT, {
    method: 'POST',
    headers,
    body: JSON.stringify(payload)
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao enviar o relatório.');
        });
      }
      return response.json().catch(() => ({}));
    })
    .then(() => {
      showToast('Relatório mensal enviado com sucesso!');
      appendBotMessage(
        'Relatório de <strong>' +
          escapeHtml(monthLabel) +
          '</strong> enviado automaticamente para <strong>' +
          escapeHtml(email) +
          '</strong>. <button type="button" class="inline-link" onclick="previewEmailTemplate()">Ver layout completo</button>',
        true
      );
      handleMonthlyEmailSuccess(email);
    })
    .catch((error) => {
      console.error('[analise:send-monthly-email]', error);
      showToast('Falha ao enviar o relatório mensal.');
      appendBotMessage(
        'Nao foi possivel concluir o envio do relatório mensal. Verifique o servidor de e-mail e tente novamente.',
        true
      );
      handleMonthlyEmailError();
    });
}

function handleIndicadoresMensais() {
  if (!state.availableMonths || !state.availableMonths.length) {
    appendBotMessage('Nenhum dado mensal disponível no momento. Verifique se a planilha de qualidade possui registros validos.', true);
    return;
  }
  const months = state.availableMonths.slice(0, 18);
  const content = createMessage('bot');
  const intro = document.createElement('p');
  intro.innerHTML = 'Selecione um mês para visualizar a analise consolidada dos fornecedores:';
  content.appendChild(intro);

  const actions = document.createElement('div');
  actions.className = 'message-actions';
  months.forEach((monthKey) => {
    const button = document.createElement('button');
    button.className = 'message-action-btn';
    button.textContent = formatMonthLabel(monthKey);
    button.addEventListener('click', () => renderMonthlyAnalysis(monthKey));
    actions.appendChild(button);
  });
  content.appendChild(actions);

  if (state.availableMonths.length > months.length) {
    const hint = document.createElement('p');
    hint.className = 'message-hint';
    hint.textContent = 'Exibindo os 18 meses mais recentes. Utilize os arquivos de origem para consultar periodos anteriores.';
    content.appendChild(hint);
  }
}

function renderMonthlyAnalysis(monthKey) {
  const monthEntry = state.monthlySummary && typeof state.monthlySummary.get === 'function' ? state.monthlySummary.get(monthKey) : null;
  if (!monthEntry) {
    appendBotMessage('Dados não encontrados para o periodo selecionado. Atualize os arquivos de origem e tente novamente.', true);
    return;
  }

  if (!monthEntry.suppliers || !monthEntry.suppliers.size) {
    appendBotMessage(
      'Nenhuma medição disponível para ' + escapeHtml(formatMonthLabel(monthKey)) + '. Atualize as planilhas e tente novamente.',
      true
    );
    return;
  }

  const monthLabel = formatMonthLabel(monthKey);
  const supplierSummaries = Array.from(monthEntry.suppliers.values()).map((entry) => ({
    key: entry.key,
    name: entry.name,
    code: entry.code,
    status: entry.status,
    sum: entry.sum,
    count: entry.count,
    avg: entry.count ? roundValue(entry.sum / entry.count) : null
  }));

  supplierSummaries.sort((a, b) => a.name.localeCompare(b.name || '', 'pt-BR', { sensitivity: 'accent' }));

  const totalSuppliers = supplierSummaries.length;
  const totalSamples = monthEntry.totalCount;
  const globalAverage = totalSamples ? roundValue(monthEntry.totalSum / totalSamples) : null;
  const reprovados = supplierSummaries
    .filter((item) => item.avg !== null && item.avg < 70)
    .sort((a, b) => (a.avg ?? 101) - (b.avg ?? 101));
  const emAtencao = supplierSummaries
    .filter((item) => item.avg !== null && item.avg >= 70 && item.avg <= 75)
    .sort((a, b) => (a.avg ?? 101) - (b.avg ?? 101));
  const aprovados = supplierSummaries.filter((item) => item.avg !== null && item.avg > 75);
  const excelencia = supplierSummaries
    .filter((item) => item.avg !== null && item.avg >= 85)
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 5);

  state.selectedMonthKey = monthKey;
  state.lastMonthlyNarrativeHtml = '';
  state.lastMonthlyEmailHtml = '';
  state.monthlySnapshot = {
    monthKey,
    monthLabel,
    generatedAt: new Date().toISOString(),
    totals: {
      totalSuppliers,
      totalSamples,
      globalAverage
    },
    distribution: {
      aprovados: aprovados.length,
      emAtencao: emAtencao.length,
      reprovados: reprovados.length
    },
    reprovados: serializeMonthlyItems(reprovados, 20),
    emAtencao: serializeMonthlyItems(emAtencao, 20),
    excelencia: serializeMonthlyItems(excelencia, 10),
    panorama: serializeMonthlyItems(supplierSummaries, 25)
  };
  clearMonthlyEmailPrompt();

  const container = createMessage('bot');
  const summaryCard = document.createElement('div');
  summaryCard.className = 'monthly-summary';
  summaryCard.innerHTML =
    '<h3>Indicadores de ' +
    escapeHtml(monthLabel) +
    '</h3>' +
    '<p><strong>Média global IQF:</strong> ' +
    (globalAverage !== null ? formatScoreValue(globalAverage) : 'N/D') +
    '</p>' +
    '<p><strong>Avaliações consideradas:</strong> ' +
    totalSamples +
    ' • <strong>Fornecedores avaliados:</strong> ' +
    totalSuppliers +
    '</p>' +
    '<p><strong>Distribuição:</strong> Aprovados ' +
    aprovados.length +
    ' | Em atenção ' +
    emAtencao.length +
    ' | Reprovados ' +
    reprovados.length +
    '</p>';
  container.appendChild(summaryCard);

  const buildListSection = (title, items, emptyMessage, cssClass) => {
    const section = document.createElement('div');
    section.className = 'monthly-subsection ' + (cssClass || '');
    section.innerHTML = '<h4>' + title + '</h4>';
    if (!items.length) {
      const empty = document.createElement('p');
      empty.className = 'empty';
      empty.textContent = emptyMessage;
      section.appendChild(empty);
      return section;
    }
    const ul = document.createElement('ul');
    items.forEach((item) => {
      const li = document.createElement('li');
      const statusTag =
        item.status === 'Reprovado'
          ? '❌'
          : item.status === 'Homologado'
          ? '🏆'
          : item.status === 'Pendente'
          ? '⏳'
          : '⚙️';
      li.textContent =
        statusTag +
        ' ' +
        item.name +
        ' — média ' +
        (item.avg !== null ? formatScoreValue(item.avg) : 'N/D') +
        ' (' +
        item.count +
        ' avaliações)';
      ul.appendChild(li);
    });
    section.appendChild(ul);
    return section;
  };

  container.appendChild(
    buildListSection(
      'Fornecedores reprovados (IQF < 70)',
      reprovados.slice(0, 10),
      'Nenhum fornecedor reprovado neste periodo.',
      'monthly-subsection-critical'
    )
  );
  if (reprovados.length > 10) {
    const note = document.createElement('p');
    note.className = 'message-hint';
    note.textContent = 'Exibindo 10 de ' + reprovados.length + ' fornecedores reprovados.';
    container.appendChild(note);
  }

  container.appendChild(
    buildListSection(
      'Fornecedores em atenção (70 ≤ IQF ≤ 75)',
      emAtencao.slice(0, 10),
      'Nenhum fornecedor em estado de atenção neste período.',
      'monthly-subsection-warning'
    )
  );
  if (emAtencao.length > 10) {
    const note = document.createElement('p');
    note.className = 'message-hint';
    note.textContent = 'Exibindo 10 de ' + emAtencao.length + ' fornecedores em atenção.';
    container.appendChild(note);
  }

  container.appendChild(
    buildListSection(
      'Destaques de excelência (IQF ≥ 85)',
      excelencia,
      'Nenhum fornecedor atingiu o patamar de excelência neste mês.',
      'monthly-subsection-success'
    )
  );

  const aiCard = document.createElement('div');
  aiCard.className = 'feedback-card';
  renderFeedbackCard(aiCard, {
    title: 'Resumo estratégico com IA',
    subtitle: 'Gerado a partir do IQF mensal',
    icon: '🧠',
    bodyHtml: '<p>Gerando análise detalhada, aguarde alguns segundos...</p>',
    hint: 'Aguarde enquanto consultamos o modelo de IA.'
  });
  container.appendChild(aiCard);

  generateMonthlyNarrative(monthKey, monthEntry, supplierSummaries, aiCard);
  renderMonthlyEmailPrompt(container, monthKey, monthLabel);
}



function serializeMonthlyItems(list, limit) {
  if (!Array.isArray(list)) {
    return [];
  }
  const max = typeof limit === 'number' ? limit : list.length;
  return list.slice(0, max).map((item) => ({
    name: item.name,
    code: item.code || null,
    status: item.status || 'Pendente',
    avg: Number.isFinite(item.avg) ? item.avg : null,
    count: item.count || 0
  }));
}

function generateMonthlyNarrative(monthKey, monthEntry, supplierSummaries, cardNode) {
  if (!cardNode) {
    return;
  }
  const apiKey = getOpenAiApiKey();
  if (!apiKey) {
    renderFeedbackCard(cardNode, {
      title: 'Resumo estratégico com IA',
      subtitle: 'Configure a chave da API para continuar',
      icon: '🔒',
      bodyHtml: '<p>Informe a chave da API OpenAI no painel de configuracoes (icone de engrenagem) para gerar a analise mensal automaticamente.</p>'
    });
    showToast('Abra o painel de configuracoes (icone de engrenagem) e cole a chave da API OpenAI para gerar o resumo mensal.');
    if (dom.apiKeyInput) {
      dom.apiKeyInput.focus();
    }
    state.lastMonthlyNarrativeHtml = '';
    refreshMonthlyEmailPromptState();
    return;
  }
  const prompt = buildMonthlyPrompt(monthKey, monthEntry, supplierSummaries);
  fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + apiKey
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: 'Você é um especialista em performance de fornecedores.' },
        { role: 'user', content: prompt }
      ]
    })
  })
    .then((response) => {
      if (!response.ok) {
        return response.text().then((text) => {
          throw new Error(text || 'Falha ao consultar a API');
        });
      }
      return response.json();
    })
    .then((payload) => {
      const answer = payload?.choices?.[0]?.message?.content?.trim();
      if (!answer) {
        throw new Error('Resposta vazia da API');
      }
      const formatted = formatFeedback(answer);
      if (state.selectedMonthKey === monthKey) {
        state.lastMonthlyNarrativeHtml = formatted;
        state.lastMonthlyEmailHtml = '';
        refreshMonthlyEmailPromptState();
      }
      renderFeedbackCard(cardNode, {
        title: 'Resumo estratégico com IA',
        subtitle: 'Complementado com IA (GPT-4o mini)',
        icon: '🤖',
        bodyHtml: formatted
      });
    })
    .catch((error) => {
      console.error('[analise:monthly-gpt]', error);
      renderFeedbackCard(cardNode, {
        title: 'Resumo estratégico com IA',
        subtitle: 'Nao foi possivel atualizar agora',
        icon: '⚠️',
        bodyHtml: '<p>Não foi possível gerar a análise neste momento. Tente novamente em instantes ou verifique a chave da API.</p>'
      });
      if (state.selectedMonthKey === monthKey) {
        state.lastMonthlyNarrativeHtml = '';
        refreshMonthlyEmailPromptState();
      }
    });
}

function buildMonthlyPrompt(monthKey, monthEntry, supplierSummaries) {
  const monthLabel = formatMonthLabel(monthKey);
  const totalEvaluations = monthEntry.totalCount;
  const globalAverage = totalEvaluations ? roundValue(monthEntry.totalSum / totalEvaluations) : null;
  const reprovados = supplierSummaries.filter((item) => item.avg !== null && item.avg < 70);
  const emAtencao = supplierSummaries.filter((item) => item.avg !== null && item.avg >= 70 && item.avg <= 75);
  const aprovados = supplierSummaries.filter((item) => item.avg !== null && item.avg > 75);
  const excelencia = supplierSummaries
    .filter((item) => item.avg !== null && item.avg >= 85)
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 10);

  const reprovLines = reprovados.length
    ? reprovados
        .slice(0, 20)
        .map((item) => '- ' + item.name + ' | média ' + formatScoreValue(item.avg) + ' | avaliações ' + item.count)
        .join('\n')
    : '- Nenhum fornecedor reprovado no período.';

  const atencaoLines = emAtencao.length
    ? emAtencao
        .slice(0, 20)
        .map((item) => '- ' + item.name + ' | média ' + formatScoreValue(item.avg) + ' | avaliações ' + item.count)
        .join('\n')
    : '- Nenhum fornecedor em atenção.';

  const excelenciaLines = excelencia.length
    ? excelencia
        .map((item) => '- ' + item.name + ' | média ' + formatScoreValue(item.avg) + ' | avaliações ' + item.count)
        .join('\n')
    : '- Nenhum destaque de excelência no periodo.';

  const baseSnapshot = supplierSummaries.slice(0, 40).map((item) => {
    if (item.avg === null) {
      return '- ' + item.name + ': sem medições registradas';
    }
    return '- ' + item.name + ': média ' + formatScoreValue(item.avg) + ' (' + item.count + ' avaliações)';
  });
  if (supplierSummaries.length > baseSnapshot.length) {
    baseSnapshot.push('- ... lista truncada para manter o prompt enxuto.');
  }

  return [
    'Analise mensal de ' + monthLabel + ' considerando ' + totalEvaluations + ' avaliações -válidas.',
    'IQF médio global: ' + (globalAverage !== null ? formatScoreValue(globalAverage) : 'N/D') + '.',
    'Distribuição por status: aprovados ' + aprovados.length + ', atenção ' + emAtencao.length + ', reprovados ' + reprovados.length + '.',
    'Destaques de excelência (>=85):\n' + excelenciaLines,
    'Fornecedores reprovados (média < 70):\n' + reprovLines,
    'Fornecedores em atenção (70 a 75):\n' + atencaoLines,
    'Panorama resumido por fornecedor:\n' + baseSnapshot.join('\n'),
    'Estruture a resposta exatamente nas seções: Visão Geral, Pontos de Atenção, Fornecedores Reprovados, Ações Imediatas, Conclusão, Alertas Prioritários.',
    'Na seção "Ações Imediatas" utilize exatamente o texto: "Envio de notificação via e-mail aos fornecedores reprovados no IQF mensal e, em caso de reincidência, abertura de RAC para analise, tratativas e possível suspensão do fornecedor."',
    'Se alguma seção nao tiver informações relevantes, escreva "Nenhum registro" de forma direta.',
    'Em "Fornecedores Reprovados" detalhe impactos e solicitação de plano de ação.',
    'Na seção "Alertas Prioritários" produza até três bullets iniciando com negrito destacando riscos críticos ou oportunidades urgentes.',
    'Use tom executivo, direto e orientado à decisão, evitando repetir dados literalmente sem interpretação. Deixando tudo organizado e bem feito.'
  ].join('\n');
}


function handleContactBase() {
  const message = [
      '<p><strong>Responsáveis pelas Bases:</strong></p>',
      '<p><strong>1️⃣ FILIAL MACAÉ:</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Naiara Galdino | Pryscila.</p>',
      '<p><strong>2️⃣ FILIAL PERNAMBUCO: </strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Suzan Bezerril.</p>',
      '<p><strong>3️⃣ FILIAL PARACURU:</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Iran Victor.</p>',
      '<p><strong>4️⃣ FILIAL FAFEN ( BA E SE ):</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Jennyfer, Gilberto, Iran, Pryscila.</p>',
      '<p><strong>5️⃣ FILIAL SÃO PAULO:</strong></p>',
      '<p><strong>Compradores:</strong></p>',
      '<p>Gilberto Trajano.</p>'

  ].join('');
   appendBotMessage(message, true);
}

function handleProcedimento() {
  const message = [
    '<p><strong>🧠 Procedimento Engeman para Avaliação de Fornecedores</strong></p>',
    '<p><strong>1️⃣ PG.SM.01 - Aquisição:</strong> Define o fluxo de compra, cotação, analise e aprovação dos materiais e serviços.</p>',
    '<p><strong>2️⃣ PG.SM.02 - Avaliação de Fornecedores:</strong> Critérios de desempenho (IQF), RACs e homologações.</p>',
    '<p><strong>3️⃣ PG.SM.03 - Almoxarifado:</strong> Inspeções, controle e tratamento de nao conformidades.</p>',
    '<p><strong>🔎 Observações Importantes</strong></p>',
    '<p><strong>- As analises são feitas mensalmente com base nas ocorrências.</strong></p>',
    '<p><strong>- A nota IQF varia de 0 a 100.</strong></p>',
    '<p><strong>- Fornecedores com IQF abaixo de 70 são REPROVADOS.</strong></p>',
    '<p><strong>- Objetivo: garantir o padrão Engeman e um relacionamento com parceiros confiáveis.</strong></p>',
    '<p><strong>📌 Dúvidas? Consulte o Setor de Suprimentos.</strong></p>'
  ].join('');
  appendBotMessage(message, true);
}


function openSettings() {
  if (dom.settingsOverlay) {
    dom.settingsOverlay.classList.add('open');
    dom.settingsOverlay.setAttribute('aria-hidden', 'false');
  }
}

function closeSettings() {
  if (dom.settingsOverlay) {
    dom.settingsOverlay.classList.remove('open');
    dom.settingsOverlay.setAttribute('aria-hidden', 'true');
  }
}

function saveSettings() {
  if (dom.settingsEmailSubject) {
    state.emailSubjectTemplate = dom.settingsEmailSubject.value.trim() || state.emailSubjectTemplate;
    localStorage.setItem('analise-email-subject', state.emailSubjectTemplate);
  }
  showToast('Configurações salvas.');
  closeSettings();
}

function clearSettings() {
  state.emailSubjectTemplate = 'Feedback IQF - {{fornecedor}}';
  if (dom.settingsEmailSubject) {
    dom.settingsEmailSubject.value = state.emailSubjectTemplate;
  }
  localStorage.removeItem('analise-email-subject');
  showToast('Configurações removidas.');
}

function createMessage(type) {
  const wrapper = document.createElement('div');
  wrapper.className = type === 'user' ? 'message user-message' : 'message bot-message';
  const content = document.createElement('div');
  content.className = 'message-content';
  wrapper.appendChild(content);
  if (dom.messages) {
    dom.messages.appendChild(wrapper);
    dom.messages.scrollTop = dom.messages.scrollHeight;
  }
  return content;
}

function appendBotMessage(text, allowHtml) {
  const content = createMessage('bot');
  if (allowHtml) {
    content.innerHTML = text;
  } else {
    content.textContent = text;
  }
  return content;
}

function appendUserMessage(text) {
  const content = createMessage('user');
  content.textContent = text;
  return content;
}

function updateMessage(node, text, allowHtml) {
  if (!node) {
    return;
  }
  if (allowHtml) {
    node.innerHTML = text;
  } else {
    node.textContent = text;
  }
}

function renderStatusBadge(status) {
  const normalized = (status || '').toLowerCase();
  let badgeClass = 'status-em-atencao';
  if (normalized === 'homologado') {
    badgeClass = 'status-aprovado';
  } else if (normalized === 'reprovado') {
    badgeClass = 'status-reprovado';
  }
  return '<span class="status-pill ' + badgeClass + '">' + escapeHtml(status || 'Pendente') + '</span>';
}

function showToast(message) {
  if (!dom.toast) {
    return;
  }
  dom.toast.innerHTML = message;
  dom.toast.classList.add('show');
  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => {
    dom.toast?.classList.remove('show');
  }, 4200);
}

function buildSupplierId(code, normalizedName, usedIds) {
  if (code) {
    const id = 'code-' + String(code).replace(/[^a-z0-9]/gi, '');
    if (!usedIds.has(id)) {
      return id;
    }
  }
  if (normalizedName) {
    const id = 'name-' + normalizedName.replace(/[^a-z0-9]/gi, '-');
    if (!usedIds.has(id)) {
      return id;
    }
  }
  let id;
  do {
    incrementalId += 1;
    id = 'supplier-' + incrementalId;
  } while (usedIds.has(id));
  return id;
}

function safeDomId(prefix, key) {
  return prefix + '-' + String(key || '').replace(/[^a-z0-9_-]/gi, '');
}

function mapHomolog(row) {
  const normalized = normalizeKeys(row);
  const code = toId(normalized.codigo || normalized.codagente || normalized.codfornecedor || normalized.id);
  const name = safeString(normalized.nomefantasia || normalized.agente || normalized.fornecedor || normalized.nome);
  if (!code && !name) {
    return null;
  }
  return {
    code,
    name,
    score: toNumber(normalized.notahomologacao || normalized.nota),
    status: mapStatus(normalized.aprovado || normalized.status || normalized.situacao || normalized.qualifica),
    expire: toISODate(normalized.datavencimento || normalized.validade || normalized.data)
  };
}

function mapIqf(row) {
  const normalized = normalizeKeys(row);
  return {
    code: toId(normalized.codagente || normalized.codigo || normalized.codfornecedor || normalized.id),
    name: safeString(normalized.nomeagente || normalized.fornecedor || normalized.nome),
    iqf: toNumber(normalized.nota || normalized.notaiqf || normalized.iqf),
    date: toISODate(normalized.data || normalized.datamedicao || normalized.datageracao),
    document: safeString(normalized.documento || normalized.numerodocumento || normalized.doc),
    occ: safeString(normalized.observacao || normalized.ocorrencia || normalized.comentario || normalized.notaocorrencia)
  };
}

function normalizeKeys(row) {
  const result = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalized = normalizeText(key).replace(/[^a-z0-9]/g, '');
    if (normalized) {
      result[normalized] = value;
    }
  });
  return result;
}

function normalizeText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/gi, '')
    .trim()
    .toLowerCase();
}

function safeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function toId(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  return String(value).trim().replace(/\.0+$/, '');
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  const cleaned = String(value).trim().replace(/\s/g, '').replace(/\.(?=\d{3}\b)/g, '').replace(',', '.');
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function toISODate(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (value instanceof Date && !Number.isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    const excelEpoch = Date.parse('1899-12-30');
    const millis = excelEpoch + value * 86400000;
    return new Date(millis).toISOString().slice(0, 10);
  }
  const match = String(value).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    const day = match[1].padStart(2, '0');
    const month = match[2].padStart(2, '0');
    const year = match[3].length === 2 ? '20' + match[3] : match[3];
    return year + '-' + month + '-' + day;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }
  return null;
}

function roundValue(value) {
  return Number.isFinite(value) ? Number(value.toFixed(2)) : value;
}

function formatMonthLabel(key) {
  if (!key) {
    return '--';
  }
  const parts = String(key).split('-');
  if (parts.length !== 2) {
    return key;
  }
  return parts[1] + '/' + parts[0];
}

function formatDate(iso) {
  if (!iso) {
    return null;
  }
  const parts = iso.split('-');
  if (parts.length !== 3) {
    return null;
  }
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

function mapStatus(value) {
  const normalized = normalizeText(value);
  if (['s', 'sim', 'homologado', 'aprovado', 'ativo', 'qualificado'].includes(normalized)) {
    return 'Homologado';
  }
  if (['n', 'nao', 'reprovado', 'bloqueado'].includes(normalized)) {
    return 'Reprovado';
  }
  return 'Pendente';
}

function deriveStatus(baseStatus, iqfScore, homologScore) {
  const belowThreshold = (score) => score !== null && score < 70;
  if (belowThreshold(iqfScore) || belowThreshold(homologScore)) {
    return 'Reprovado';
  }
  if (baseStatus === 'Homologado') {
    return 'Homologado';
  }
  if (baseStatus === 'Reprovado') {
    return 'Reprovado';
  }
  return 'Pendente';
}

function severityLevel(text) {
  const normalized = normalizeText(text);
  if (!normalized) {
    return 'low';
  }
  if (/nao\s+efetuou|falha|critico|grave/.test(normalized)) {
    return 'critical';
  }
  if (/atraso|problema|irregular|bloqueio/.test(normalized)) {
    return 'high';
  }
  if (/pendente|ajuste|analise|monitor/.test(normalized)) {
    return 'medium';
  }
  return 'low';
}

function escapeHtml(value) {
  return String(value || '').replace(/[&<>"']/g, (match) =>
    ({
      '&': '&',
      '<': '<',
      '>': '>',
      '"': '"',
      "'": '&#39;'
    }[match])
  );
}

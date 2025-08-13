/**
 *
 * ARQUIVO DE CONFIGURAÇÃO GLOBAL
 *
 * Este arquivo centraliza todas as variáveis de configuração para os scripts do projeto.
 * Ele NÃO deve ser executado diretamente. Ele apenas define as variáveis para os outros scripts.
 */

move file to src/Configuracoes.gs
const ID_PLANILHA_PRINCIPAL = '1nINa_O3EFr9GugHc3SIch2JfeMkWGgIJzVdKTeWSSTQ';
const NOME_ABA_DADOS_GCR = 'bd_gcr_script';
const NOME_ABA_DESTINO_CARTEIRA = 'bd_carteira_script';
const NOME_ABA_CONFIG_CONSULTORES = 'Config_Consultores';
const NOME_ABA_PEP = 'base_pep';

// =================== CONFIGURAÇÕES DE ANÁLISE DE DADOS ===================
const NOME_ABA_ANALISE_ESTATISTICA = 'Analise_Estatistica';
const NOME_ABA_ANALISE_DIAGNOSTICA = 'Analise_Diagnostica';

// =================== CONFIGURAÇÕES DE PROCESSAMENTO DE DADOS (ETL) ===================
// --- Origem GCR (Formulário) ---
const ID_PLANILHA_ORIGEM_GCR = '1hq83VYnoFvUpEpvUZ1u5qrmi4FfgYNvWXEgRHWE4RoI';
const NOME_ABA_ORIGEM_GCR = 'Respostas ao formulário 1';

// --- Origem Carteira ---
const ID_PLANILHA_ORIGEM_CARTEIRA = "15YHghamMti938P8hyVpuA5G3xKZnnBKBkOoAk1c_fmM";
const NOME_ABA_ORIGEM_CARTEIRA = "Carteira";

// =================== CONFIGURAÇÕES DE ALERTA DE CHURN ===================

// --- E-mail (Alerta Geral para gestores) ---
const EMAIL_DESTINATARIO_ALERTA = 'thales.bregantin@mercadolivre.com'; 

// --- Workchat (Fallback/Padrão) ---
// Webhook para um canal geral, caso o consultor não seja encontrado na aba de configuração.
const DEFAULT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/_2HQi8AAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=vhG69pjz08nM7XGP9zKKZwNZ8vKZjrS0KgqmR9Afrw8';

// --- Gatilho de Risco ---
// Define qual evento específico deve disparar o alerta.
const GATILHO_DE_RISCO = '04. Gestão de pausas';
const CAUSAL_DE_RISCO = 'Solicitação de pausa repentina';

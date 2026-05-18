/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useRef, type RefObject, type FormEvent } from 'react';
import { Routes, Route, useNavigate, useLocation } from 'react-router-dom';
import { 
  CheckCircle2, XCircle, ChevronLeft, ChevronRight, RefreshCcw, 
  X, Info, Download, BookOpen, ShieldCheck, Search,
  TrendingUp, TrendingDown, Activity, Settings, LayoutGrid,
  Clock, Bell, Triangle, LayoutDashboard, Layers, CreditCard, Users as UsersIcon, Plus, Trash2, Edit2, Save, LogOut
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, 
  ResponsiveContainer, LineChart, Line 
} from 'recharts';
import type { Section, Metric, MetricHistory, User, AccessLevel } from './types';
import { useDataPersistence } from './services/dataPersistence';

const INITIAL_DATA: Section[] = [
  {
    id: 'sec-1',
    title: "Qualidade operacional",
    metrics: [
      { 
        id: '1', title: "Status separável", value: 12, status: 'critical', lastUpdate: "13/05/2026, 20:37:06",
        objective: "Garantir que todos os itens marcados como separáveis possuam status válido no sistema para processamento.",
        rules: [
          "Regra 1: O item deve estar em uma área de picking ativa.",
          "Regra 2: O status do LPN deve ser 'Disponível' ou 'Reservado'."
        ],
        query: "SELECT COUNT(*) FROM inventory WHERE separable = true AND status NOT IN ('Available', 'Reserved')",
        history: Array.from({ length: 12 }, (_, i) => ({ date: `${10+i}:00`, value: Math.floor(Math.random() * 20) })),
        details: [
          { id: 'dt1', posicao: 'P-01-A', item: 'MOUSE LOGITECH G502', validade: 'N/A', lote: 'L1', quantidade: 5, motivo: 'Status Inválido' },
          { id: 'dt2', posicao: 'P-01-B', item: 'TECLADO RAZER', validade: 'N/A', lote: 'L2', quantidade: 2, motivo: 'Status Inválido' },
        ]
      },
      { 
        id: '2', title: "RF's desbloqueados", value: 405, status: 'critical', lastUpdate: "13/05/2026, 20:29:06", 
        objective: "Identificar o recebimento de inbound realizado utilizando RF que não esteja bloqueado.",
        rules: [
          "Regra 1: Área: Estar entre RSTG, SSTG Será considerado Não Auditável.",
          "Regra 2: Área: Estar entre RCKTAERBLO... Status: Diferente de Bloqueado e Recall. Será considerado incorreto."
        ],
        query: "SELECT COUNT(*) FROM rf_devices WHERE unlocked = true AND status != 'Blocked'",
        history: Array.from({ length: 12 }, (_, i) => ({ date: `${10+i}:00`, value: 350 + Math.floor(Math.random() * 100) })),
        details: [
          { id: 'd1', posicao: 'A-12-01', item: 'RECEPTOR RF-900', validade: 'N/A', lote: 'LT9920', quantidade: 150, motivo: 'Bloqueio de Qualidade' },
        ]
      },
      { id: '3', title: "Recebimento de correlatos", value: 25, status: 'critical', lastUpdate: "13/05/2026, 20:30:06", query: "SELECT * FROM receipts WHERE family = 'correlatos'", history: Array.from({ length: 10 }, (_, i) => ({ date: `${10+i}:00`, value: 10 + i * 2 })) },
      { id: '4', title: "Invoice manual", value: 18, status: 'critical', lastUpdate: "13/05/2026, 20:33:06", query: "SELECT * FROM invoices WHERE manual = true" },
      { id: '5', title: "Pendência conferência", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:35:00", query: "SELECT * FROM pending_check" },
    ]
  },
  {
    id: 'sec-2',
    title: "Validação sistêmica",
    metrics: [
      { id: '6', title: "Status x Área", value: 333, status: 'critical', lastUpdate: "13/05/2026, 20:29:53", query: "SELECT * FROM sys_val WHERE type = 'status_area'" },
      { id: '7', title: "Família x Área", value: 2308, status: 'critical', lastUpdate: "13/05/2026, 20:30:08", query: "SELECT * FROM sys_val WHERE type = 'family_area'" },
      { id: '8', title: "Código verificador", value: 5542, status: 'critical', lastUpdate: "13/05/2026, 20:31:09", query: "SELECT * FROM sys_val WHERE type = 'verifier'" },
      { id: '9', title: "Flag de separável", value: 1132, status: 'critical', lastUpdate: "13/05/2026, 20:31:20", query: "SELECT * FROM sys_val WHERE type = 'separable_flag'" },
      { id: '10', title: "Divergência lote", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:31:25", query: "SELECT * FROM sys_val WHERE type = 'batch_div'" },
    ]
  },
  {
    id: 'sec-3',
    title: "Validação de saldo",
    metrics: [
      { id: '11', title: "LPN", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:31:27" },
      { id: '12', title: "SKU armazenado", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:31:29" },
      { id: '13', title: "Status x Regra Aging", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:31:31" },
      { id: '14', title: "Manufatura futura", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:36:57" },
      { id: '15', title: "Expiração", value: 0, status: 'ok', lastUpdate: "13/05/2026, 20:37:01" },
      { id: '16', title: "Quantidade de Aging", value: 525, status: 'critical', lastUpdate: "13/05/2026, 20:37:08" },
    ]
  }
];

const INITIAL_USERS: User[] = [
  { id: 'u-1', name: 'Arlen Loran', email: 'arlenloran@gmail.com', role: 'admin', lastActive: '14/05/2026, 01:05:07', status: 'Ativo' },
  { id: 'u-2', name: 'Editor Operacional', email: 'editor@empresa.com', role: 'editor', lastActive: '13/05/2026, 18:30:15', status: 'Ativo' },
  { id: 'u-3', name: 'Visualizador Geral', email: 'viewer@empresa.com', role: 'viewer', lastActive: '14/05/2026, 00:15:44', status: 'Ativo' },
];

function useScrollIndicator(ref: RefObject<HTMLDivElement | null>) {
  const [scrollInfo, setScrollInfo] = useState({ percentage: 0, ratio: 0, isScrollable: false });

  const updateScroll = () => {
    if (ref.current) {
      const { scrollLeft, scrollWidth, clientWidth } = ref.current;
      const isScrollable = scrollWidth > clientWidth + 1;
      const percentage = isScrollable ? (scrollLeft / (scrollWidth - clientWidth)) * 100 : 0;
      const ratio = isScrollable ? clientWidth / scrollWidth : 1;
      setScrollInfo({ percentage, ratio, isScrollable });
    }
  };

  useEffect(() => {
    const el = ref.current;
    if (!el) return;

    updateScroll();
    
    const observer = new ResizeObserver(updateScroll);
    observer.observe(el);
    
    el.addEventListener('scroll', updateScroll);
    
    return () => {
      observer.disconnect();
      el.removeEventListener('scroll', updateScroll);
    };
  }, [ref]);

  return scrollInfo;
}

interface MetricCardProps {
  metric: Metric;
  onClick?: (metric: Metric) => void;
  isWarRoom?: boolean;
  countdown?: number;
  key?: string;
}

function MiniSparkline({ data, color }: { data: MetricHistory[], color: string }) {
  const chartData = data.map((item, i) => ({ value: item.value, index: i }));
  return (
    <div className="w-full h-8 mt-2 overflow-hidden opacity-50 group-hover:opacity-100 transition-opacity duration-500">
      <ResponsiveContainer width="100%" height="100%">
        <AreaChart data={chartData}>
          <defs>
            <linearGradient id={`gradient-${color}`} x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={color} stopOpacity={0.3}/>
              <stop offset="95%" stopColor={color} stopOpacity={0}/>
            </linearGradient>
          </defs>
          <Area 
            type="monotone" 
            dataKey="value" 
            stroke={color} 
            strokeWidth={1.5} 
            fillOpacity={1} 
            fill={`url(#gradient-${color})`} 
            isAnimationActive={false}
          />
        </AreaChart>
      </ResponsiveContainer>
    </div>
  );
}

function MetricCard({ metric, onClick, isWarRoom, countdown }: MetricCardProps) {
  const hasHistory = metric.history && metric.history.length > 1;
  let trend: { value: number; isUp: boolean; text: string } | null = null;
  
  if (hasHistory) {
    const lastValue = metric.history![metric.history!.length - 1].value;
    const prevValue = metric.history![metric.history!.length - 2].value;
    const diff = lastValue - prevValue;
    if (prevValue !== 0) {
      const percent = Math.abs((diff / prevValue) * 100).toFixed(1);
      trend = { value: Math.abs(diff), isUp: diff > 0, text: `${diff > 0 ? '+' : ''}${percent}%` };
    } else if (diff !== 0) {
      trend = { value: Math.abs(diff), isUp: diff > 0, text: `${diff > 0 ? '+' : ''}${diff}` };
    }
  }

  return (
    <motion.div
      layout
      initial={{ opacity: 0, scale: 0.95 }}
      animate={{ opacity: 1, scale: 1 }}
      onClick={() => onClick?.(metric)}
      whileHover={{ 
        y: -4, 
        boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1)",
        transition: { duration: 0.2 } 
      }}
      className={`rounded-xl shadow-sm p-3 min-w-[170px] h-full flex flex-col items-center justify-between border transition-all duration-500 cursor-pointer relative overflow-hidden ${
        isWarRoom 
        ? 'bg-[#0b0e1a] border-indigo-900/30 shadow-slate-950/50 hover:border-indigo-500/50' 
        : 'bg-white border-slate-100 hover:border-slate-300'
      }`}
      id={`card-${metric.id}`}
    >
      {trend && (
        <div className={`absolute top-2 right-2 flex items-center gap-0.5 px-1.5 py-0.5 rounded-full text-[9px] font-black transition-colors duration-500 ${
          trend.isUp 
          ? (isWarRoom ? 'bg-red-950/50 text-red-500' : 'bg-red-50 text-red-600') 
          : (isWarRoom ? 'bg-emerald-950/50 text-emerald-500' : 'bg-emerald-50 text-emerald-600')
        }`}>
          {trend.isUp ? <TrendingUp className="w-2.5 h-2.5" /> : <TrendingDown className="w-2.5 h-2.5" />}
          {trend.text}
        </div>
      )}
      <header className="w-full text-center mb-1">
        <h3 className={`text-[10px] font-bold uppercase tracking-wide truncate py-2 transition-colors duration-500 ${trend ? 'pl-3 pr-12' : 'px-3'} ${isWarRoom ? 'text-slate-300' : 'text-slate-700'}`}>
          {metric.title}
        </h3>
        <div className={`h-px w-4/5 mx-auto transition-colors duration-500 ${isWarRoom ? 'bg-slate-800' : 'bg-slate-200'}`} />
      </header>
      
      <div className="flex flex-col items-center gap-3 py-1">
        <span className={`text-2xl font-black italic tracking-tighter tabular-nums transition-colors duration-500 ${isWarRoom ? 'text-white text-3xl' : 'text-slate-900'}`}>
          {metric.value}
        </span>
        
        <div className={`relative group/icon flex items-center justify-center min-h-[64px] rounded-full p-1 transition-colors duration-500 ${isWarRoom ? 'bg-slate-800/50' : 'bg-transparent'}`}>
          {metric.status === 'ok' ? (
            <motion.div 
              whileHover={{ scale: 1.1 }}
              transition={{ type: "spring", stiffness: 400, damping: 10 }}
            >
              <CheckCircle2 className="w-14 h-14 text-emerald-500" strokeWidth={1.5} />
            </motion.div>
          ) : (
            <motion.div
              animate={{ 
                scale: [1, isWarRoom ? 1.4 : 1.3, 1],
              }}
              transition={{ 
                repeat: Infinity, 
                duration: isWarRoom ? 0.6 : 0.8,
                ease: "easeInOut"
              }}
              style={{ filter: isWarRoom ? "drop-shadow(0 0 15px rgba(255, 0, 0, 0.4))" : "drop-shadow(0 0 10px rgba(204, 0, 0, 0.25))" }}
              whileHover={{ scale: 1.4 }}
            >
              <XCircle className={`w-14 h-14 ${isWarRoom ? 'text-red-500' : 'text-brand-red'}`} strokeWidth={1.5} />
            </motion.div>
          )}
        </div>
      </div>
      
      <footer className="w-full mt-auto text-center">
        <div className={`h-px w-4/5 mx-auto mb-2 transition-colors duration-500 ${isWarRoom ? 'bg-slate-800' : 'bg-slate-200'}`} />
        
        <div className="h-10 flex flex-col justify-center mb-1">
          {metric.history && metric.history.length > 2 ? (
            <MiniSparkline 
              data={metric.history} 
              color={metric.status === 'ok' ? '#10b981' : '#ef4444'} 
            />
          ) : (
            <div className="w-full h-8" /> 
          )}
        </div>

        <div className="flex items-center justify-center gap-1.5 mt-1">
          <Clock className={`w-2.5 h-2.5 ${isWarRoom ? 'text-slate-600' : 'text-slate-300'}`} />
          <span className={`text-[10px] font-semibold tracking-tighter transition-colors duration-500 ${isWarRoom ? 'text-slate-500' : 'text-slate-400'}`}>
            {metric.lastUpdate} {countdown !== undefined ? `(${countdown}s)` : ''}
          </span>
        </div>
      </footer>

      {isWarRoom && metric.status === 'critical' && (
        <div className="absolute inset-x-0 bottom-0 h-1 bg-red-500 animate-pulse" />
      )}
    </motion.div>
  );
}

function SectionContainer({ section, onCardClick, isWarRoom, metricCountdowns }: { 
  section: Section, 
  onCardClick: (metric: Metric) => void, 
  isWarRoom: boolean,
  metricCountdowns: Record<string, number>
}) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const { percentage, ratio, isScrollable } = useScrollIndicator(scrollRef);
  const scrollInterval = useRef<NodeJS.Timeout | null>(null);

  useEffect(() => {
    if (isWarRoom && isScrollable) {
      scrollInterval.current = setInterval(() => {
        if (scrollRef.current) {
          const { scrollLeft, scrollWidth, clientWidth } = scrollRef.current;
          if (scrollLeft + clientWidth >= scrollWidth - 5) {
            scrollRef.current.scrollTo({ left: 0, behavior: 'smooth' });
          } else {
            scrollRef.current.scrollBy({ left: 300, behavior: 'smooth' });
          }
        }
      }, 6000);
    } else {
      if (scrollInterval.current) clearInterval(scrollInterval.current);
    }
    return () => {
      if (scrollInterval.current) clearInterval(scrollInterval.current);
    };
  }, [isWarRoom, isScrollable]);

  const scroll = (direction: 'left' | 'right') => {
    if (scrollRef.current) {
      const amount = direction === 'left' ? -300 : 300;
      scrollRef.current.scrollBy({ left: amount, behavior: 'smooth' });
    }
  };

  return (
    <section className={`flex flex-col gap-3 transition-all duration-500 h-full ${isWarRoom ? 'scale-[1.02]' : ''}`} id={`section-${section.title.replace(/\s/g, '-')}`}>
      <div className={`py-2 px-5 rounded-xl shadow-sm inline-flex items-center w-full flex-shrink-0 transition-colors duration-500 ${isWarRoom ? 'bg-[#0b0e1a] border border-indigo-900/30' : 'bg-brand-red text-white'}`}>
        <h2 className={`text-lg font-bold tracking-tight uppercase italic ${isWarRoom ? 'text-brand-red drop-shadow-[0_0_8px_rgba(204,0,0,0.3)]' : 'text-white'}`}>
          {section.title}
        </h2>
      </div>
      
      <div className="relative group flex-grow flex flex-col">
        <div 
          ref={scrollRef}
          className="overflow-x-auto pb-4 scrollbar-hide flex gap-3 items-stretch flex-grow scroll-smooth px-1"
        >
          {section.metrics.map((metric) => (
            <MetricCard 
              key={metric.id} 
              metric={metric} 
              onClick={onCardClick} 
              isWarRoom={isWarRoom} 
              countdown={metricCountdowns[metric.id]}
            />
          ))}
        </div>
        
        <AnimatePresence>
          {isScrollable && (
            <motion.div 
              initial={{ opacity: 0, y: 5 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 5 }}
              className="flex items-center justify-between mt-1 px-1"
            >
              <button 
                onClick={() => scroll('left')}
                className={`focus:outline-none hover:scale-110 transition-all active:scale-95 ${isWarRoom ? 'text-slate-500 hover:text-white' : 'text-brand-yellow'}`}
              >
                <ChevronLeft className="w-5 h-5" strokeWidth={isWarRoom ? 2 : 4} />
              </button>
              
              <div className={`mx-4 h-1.5 flex-grow rounded-full relative overflow-hidden shadow-inner transition-colors duration-500 ${isWarRoom ? 'bg-slate-900' : 'bg-brand-yellow/10'}`}>
                <motion.div 
                  className={`absolute top-0 left-0 h-full rounded-full transition-colors duration-500 ${isWarRoom ? 'bg-brand-red shadow-[0_0_8px_rgba(239,68,68,0.6)]' : 'bg-brand-yellow'}`}
                  style={{ 
                    width: `${ratio * 100}%`, 
                    left: `${percentage * (1 - ratio)}%` 
                  }}
                  transition={{ type: 'spring', bounce: 0, duration: 0.1 }}
                />
              </div>
              
              <button 
                onClick={() => scroll('right')}
                className={`focus:outline-none hover:scale-110 transition-all active:scale-95 ${isWarRoom ? 'text-slate-500 hover:text-white' : 'text-brand-yellow'}`}
              >
                <ChevronRight className="w-5 h-5" strokeWidth={isWarRoom ? 2 : 4} />
              </button>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </section>
  );
}

function DivergenceModal({ metric, onClose }: { metric: Metric, onClose: () => void }) {
  const [activeTab, setActiveTab] = useState<'table' | 'objective' | 'rules' | 'trend'>('table');
  const [searchQuery, setSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;
  
  if (!metric) return null;

  const filteredDetails = (metric.details || []).filter(item => {
    const query = searchQuery.toLowerCase();
    return (
      item.posicao.toLowerCase().includes(query) ||
      item.item.toLowerCase().includes(query) ||
      item.validade.toLowerCase().includes(query) ||
      item.lote.toLowerCase().includes(query) ||
      item.quantidade.toString().includes(query) ||
      item.motivo.toLowerCase().includes(query)
    );
  });

  const totalPages = Math.ceil(filteredDetails.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const paginatedDetails = filteredDetails.slice(startIndex, startIndex + itemsPerPage);

  const downloadExcel = () => {
    if (!metric.details || metric.details.length === 0) return;
    
    const worksheet = XLSX.utils.json_to_sheet(metric.details.map(d => ({
      'Posição': d.posicao,
      'Item': d.item,
      'Validade': d.validade,
      'Lote': d.lote,
      'Quantidade': d.quantidade,
      'Motivo': d.motivo
    })));
    
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Divergências");
    XLSX.writeFile(workbook, `Divergencias_${metric.title.replace(/\s/g, '_')}.xlsx`);
  };

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
      onClick={onClose}
    >
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="px-6 py-4 border-b border-slate-100 flex flex-col sm:flex-row justify-between items-start sm:items-center bg-slate-50/50 gap-4">
          <div className="flex items-center gap-3">
            <div className={`w-2 h-8 rounded-full ${metric.status === 'critical' ? 'bg-brand-red' : 'bg-emerald-500'}`} />
            <div>
              <h2 className="text-lg font-black text-slate-900 uppercase tracking-tight italic">
                Detalhes: <span className="text-brand-red">{metric.title}</span>
              </h2>
              <p className="text-[10px] font-bold text-slate-500 uppercase">
                Variações Identificadas - {metric.lastUpdate}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 self-end sm:self-auto">
            <div className="flex bg-slate-200 p-1 rounded-lg">
              <button 
                onClick={() => setActiveTab('table')}
                className={`p-2 rounded-md transition-all ${activeTab === 'table' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}
                title="Visualizar Tabela"
              >
                <Info className="w-5 h-5" />
              </button>
              <button 
                onClick={() => setActiveTab('objective')}
                className={`p-2 rounded-md transition-all ${activeTab === 'objective' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}
                title="Objetivo da Métrica"
              >
                <BookOpen className="w-5 h-5" />
              </button>
              <button 
                onClick={() => setActiveTab('rules')}
                className={`p-2 rounded-md transition-all ${activeTab === 'rules' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}
                title="Regras Aplicadas"
              >
                <ShieldCheck className="w-5 h-5" />
              </button>
              <button 
                onClick={() => setActiveTab('trend')}
                className={`p-2 rounded-md transition-all ${activeTab === 'trend' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}
                title="Histórico Comparativo"
              >
                <TrendingUp className="w-5 h-5" />
              </button>
            </div>
            
            <button 
              onClick={downloadExcel}
              className="p-2 bg-emerald-500 text-white rounded-lg hover:bg-emerald-600 transition-all shadow-sm active:scale-95 disabled:opacity-50"
              title="Download Excel"
              disabled={!metric.details || metric.details.length === 0}
            >
              <Download className="w-5 h-5" />
            </button>

            <button 
              onClick={onClose}
              className="p-2 hover:bg-slate-200 rounded-full transition-colors text-slate-400 hover:text-slate-600"
            >
              <X className="w-6 h-6" />
            </button>
          </div>
        </header>

        <div className="flex-grow overflow-auto p-6">
          <AnimatePresence mode="wait">
            {activeTab === 'table' && (
              <motion.div
                key="table"
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 10 }}
                transition={{ duration: 0.2 }}
                className="space-y-4"
              >
                {metric.details && metric.details.length > 0 ? (
                  <>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input 
                        type="text" 
                        placeholder="Buscar em qualquer coluna..."
                        value={searchQuery}
                        onChange={(e) => {
                          setSearchQuery(e.target.value);
                          setCurrentPage(1);
                        }}
                        className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-slate-200 focus:border-slate-400 outline-none transition-all font-medium"
                      />
                    </div>

                    <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200">
                            <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">Posição</th>
                            <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">Item / Descrição</th>
                            <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider text-center">Validade</th>
                            <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider text-center">Lote</th>
                            <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider text-center">Qtd.</th>
                            <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">Motivo Divergência</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {paginatedDetails.length > 0 ? (
                            paginatedDetails.map((detail) => (
                              <tr key={detail.id} className="hover:bg-slate-50 transition-colors group">
                                <td className="px-4 py-4 text-sm font-bold text-slate-900 font-mono">{detail.posicao}</td>
                                <td className="px-4 py-4 text-sm font-medium text-slate-700">{detail.item}</td>
                                <td className="px-4 py-4 text-sm text-slate-600 text-center font-mono">{detail.validade}</td>
                                <td className="px-4 py-4 text-sm text-slate-600 text-center font-mono">{detail.lote}</td>
                                <td className="px-4 py-4 text-sm font-black text-brand-red text-center tabular-nums">{detail.quantidade}</td>
                                <td className="px-4 py-4 text-sm">
                                  <span className="px-2 py-1 bg-red-50 text-red-700 rounded-lg font-bold text-[10px] border border-red-100">
                                    {detail.motivo}
                                  </span>
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={6} className="px-4 py-10 text-center font-medium text-slate-500 italic">
                                Nenhum resultado encontrado para "{searchQuery}"
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>

                    {totalPages > 1 && (
                      <div className="flex items-center justify-between px-2 pt-2">
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                          Exibindo {startIndex + 1}-{Math.min(startIndex + itemsPerPage, filteredDetails.length)} de {filteredDetails.length} itens
                        </span>
                        <div className="flex gap-1">
                          <button 
                            onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                            disabled={currentPage === 1}
                            className="p-2 hover:bg-slate-100 rounded-lg disabled:opacity-30 transition-colors"
                          >
                            <ChevronLeft className="w-5 h-5" />
                          </button>
                          <div className="flex items-center gap-1 px-2">
                            {Array.from({ length: totalPages }, (_, i) => i + 1).map((page) => (
                              <button
                                key={page}
                                onClick={() => setCurrentPage(page)}
                                className={`w-8 h-8 rounded-lg text-xs font-bold transition-all ${currentPage === page ? 'bg-slate-900 text-white shadow-md' : 'text-slate-400 hover:bg-slate-100'}`}
                              >
                                {page}
                              </button>
                            ))}
                          </div>
                          <button 
                            onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                            disabled={currentPage === totalPages}
                            className="p-2 hover:bg-slate-100 rounded-lg disabled:opacity-30 transition-colors"
                          >
                            <ChevronRight className="w-5 h-5" />
                          </button>
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <div className="flex flex-col items-center justify-center py-20 text-center space-y-4">
                    <div className="w-20 h-20 rounded-full bg-emerald-50 flex items-center justify-center border-4 border-emerald-100">
                      <CheckCircle2 className="w-10 h-10 text-emerald-500" />
                    </div>
                    <div>
                      <h3 className="text-xl font-black text-slate-900 tracking-tight">SEM DIVERGÊNCIAS</h3>
                      <p className="text-slate-500 font-medium">Todos os registros para esta métrica estão em conformidade.</p>
                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {activeTab === 'objective' && (
              <motion.div
                key="objective"
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 10 }}
                transition={{ duration: 0.2 }}
                className="bg-slate-50 rounded-xl p-8 border border-slate-200"
              >
                <div className="flex items-center gap-3 mb-6">
                  <div className="p-3 bg-blue-100 text-blue-600 rounded-xl">
                    <BookOpen className="w-8 h-8" />
                  </div>
                  <h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Objetivo da Métrica</h3>
                </div>
                <p className="text-lg text-slate-700 leading-relaxed font-medium">
                  {metric.objective || "Nenhum objetivo detalhado cadastrado para esta métrica."}
                </p>
              </motion.div>
            )}

            {activeTab === 'rules' && (
              <motion.div
                key="rules"
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 10 }}
                transition={{ duration: 0.2 }}
                className="space-y-4"
              >
                <div className="flex items-center gap-3 mb-4">
                  <div className="p-3 bg-amber-100 text-amber-600 rounded-xl">
                    <ShieldCheck className="w-8 h-8" />
                  </div>
                  <h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Regras de Validação</h3>
                </div>
                
                {metric.rules && metric.rules.length > 0 ? (
                  <div className="grid gap-3">
                    {metric.rules.map((rule, idx) => (
                      <div key={idx} className="bg-white border border-slate-200 p-4 rounded-xl shadow-sm border-l-4 border-l-amber-500">
                        <p className="text-slate-700 font-bold whitespace-pre-line text-sm leading-relaxed">
                          {rule}
                        </p>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="text-slate-500 italic p-8 bg-slate-50 rounded-xl border border-slate-200 text-center">
                    Nenhuma regra de validação cadastrada para esta métrica.
                  </p>
                )}
              </motion.div>
            )}
            {activeTab === 'trend' && (
              <motion.div
                key="trend"
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 10 }}
                transition={{ duration: 0.2 }}
                className="space-y-6"
              >
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-3">
                    <div className="p-3 bg-indigo-100 text-indigo-600 rounded-xl">
                      <Activity className="w-8 h-8" />
                    </div>
                    <h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Análise de Tendência</h3>
                  </div>
                  <div className="text-right">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Valor Atual</p>
                    <p className="text-3xl font-black text-slate-900 tabular-nums">{metric.value}</p>
                  </div>
                </div>

                {metric.history && metric.history.length > 0 ? (
                  <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm h-[350px] relative">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={metric.history.map((val, idx) => ({ value: val, time: `T-${metric.history!.length - idx - 1}h` }))}>
                        <defs>
                          <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#cc0000" stopOpacity={0.1}/>
                            <stop offset="95%" stopColor="#cc0000" stopOpacity={0}/>
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis 
                          dataKey="time" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fontSize: 10, fontWeight: 700, fill: '#94a3b8' }}
                          dy={10}
                        />
                        <YAxis 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fontSize: 10, fontWeight: 700, fill: '#94a3b8' }}
                        />
                        <Tooltip 
                          contentStyle={{ 
                            borderRadius: '12px', 
                            border: 'none', 
                            boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                            fontSize: '12px',
                            fontWeight: 'bold'
                          }}
                        />
                        <Area 
                          type="monotone" 
                          dataKey="value" 
                          stroke="#cc0000" 
                          strokeWidth={4} 
                          fillOpacity={1} 
                          fill="url(#colorValue)" 
                          animationDuration={1500}
                        />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-12 text-center">
                    <p className="text-slate-500 font-bold italic">Nenhum dado histórico disponível para esta métrica ainda.</p>
                  </div>
                )}

                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Média</p>
                    <p className="text-lg font-black text-slate-800">
                      {metric.history ? (metric.history.reduce((acc, curr) => acc + curr.value, 0) / metric.history.length).toFixed(1) : '0'}
                    </p>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Pico</p>
                    <p className="text-lg font-black text-slate-800">
                      {metric.history ? Math.max(...metric.history.map(h => h.value)) : '0'}
                    </p>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Status</p>
                    <p className={`text-lg font-black ${metric.status === 'critical' ? 'text-brand-red' : 'text-emerald-500'}`}>
                      {metric.status === 'critical' ? 'CRÍTICO' : 'ESTÁVEL'}
                    </p>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Meta</p>
                    <p className="text-lg font-black text-slate-800">0</p>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        <footer className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex justify-between items-center">
          <div className="flex items-center gap-2 text-slate-400">
            <Info className="w-4 h-4" />
            <span className="text-[10px] font-bold uppercase tracking-widest">Informações em tempo real do ERP/WMS</span>
          </div>
          <button 
            onClick={onClose}
            className="px-6 py-2 bg-slate-900 text-white rounded-lg font-black text-xs uppercase tracking-widest hover:bg-slate-800 transition-all active:scale-95 shadow-md"
          >
            Fechar Painel
          </button>
        </footer>
      </motion.div>
    </motion.div>
  );
}

function Sidebar({ currentView, setView, isWarRoom }: { currentView: string, setView: (v: string) => void, isWarRoom: boolean }) {
  const menus = [
    { id: 'dashboard', icon: LayoutDashboard, label: 'Dashboard' },
    { id: 'divisions', icon: Layers, label: 'Divisões' },
    { id: 'cards', icon: CreditCard, label: 'Cards' },
    { id: 'users', icon: UsersIcon, label: 'Usuários' },
  ];

  return (
    <aside className={`w-20 md:w-64 h-screen fixed left-0 top-0 z-[100] border-r transition-all duration-700 flex flex-col ${
      isWarRoom ? 'bg-[#0b0e1a] border-indigo-900/40 text-white' : 'bg-white border-slate-200 text-slate-900'
    }`}>
      <div className="p-6 mb-10 overflow-hidden">
        <div className="flex items-center gap-3">
          <Activity className="w-8 h-8 text-brand-red flex-shrink-0" />
          <h1 className="text-xl font-black italic tracking-tighter uppercase leading-tight truncate hidden md:block">
            Monitor <span className="text-brand-red">System</span>
          </h1>
        </div>
      </div>

      <nav className="flex-grow px-3 space-y-2">
        {menus.map((menu) => (
          <button
            key={menu.id}
            onClick={() => setView(menu.id)}
            className={`w-full flex items-center gap-4 px-4 py-4 rounded-2xl transition-all group relative ${
              currentView === menu.id 
              ? 'bg-brand-red text-white shadow-[0_10px_20px_rgba(204,0,0,0.3)]' 
              : `hover:bg-slate-100 ${isWarRoom ? 'hover:bg-white/5 text-slate-400' : 'text-slate-500'}`
            }`}
          >
            <menu.icon className={`w-6 h-6 flex-shrink-0 transition-transform ${currentView === menu.id ? '' : 'group-hover:scale-110'}`} />
            <span className="font-black text-xs uppercase tracking-widest truncate hidden md:block">{menu.label}</span>
            {currentView === menu.id && (
              <motion.div layoutId="activeNav" className="absolute right-0 top-1/2 -translate-y-1/2 w-1 h-8 bg-white rounded-l-full" />
            )}
          </button>
        ))}
      </nav>

      <div className="p-6 mt-auto">
        <button className="w-full flex items-center gap-4 px-4 py-4 rounded-2xl text-slate-500 hover:text-brand-red transition-colors group">
          <LogOut className="w-6 h-6 group-hover:-translate-x-1 transition-transform" />
          <span className="font-black text-xs uppercase tracking-widest hidden md:block">Logout</span>
        </button>
      </div>
    </aside>
  );
}

function SectionManager({ data, setData, isWarRoom, updateSection, deleteSection }: { 
  data: Section[], 
  setData: (d: Section[] | ((prev: Section[]) => Section[])) => void, 
  isWarRoom: boolean,
  updateSection: (s: Section) => Promise<void>,
  deleteSection: (s: Section) => Promise<void>
}) {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [newName, setNewName] = useState('');
  const [isModalOpen, setIsModalOpen] = useState(false);

  const saveSectionLocal = async (e: FormEvent) => {
    e.preventDefault();
    if (!newName) return;
    const newSection: Section = { id: `sec-${Date.now()}`, title: newName, metrics: [] };
    await updateSection(newSection);
    setIsModalOpen(false);
    setNewName('');
  };

  const syncSectionName = async (id: string) => {
    const section = data.find(s => s.id === id);
    if (section) {
      await updateSection({ ...section, title: newName });
    }
    setEditingId(null);
    setNewName('');
  };

  const removeSection = async (section: Section) => {
    if (confirm(`Excluir a divisão "${section.title}"?`)) {
      await deleteSection(section);
    }
  };

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <header className="flex justify-between items-end mb-10">
        <div>
          <h2 className={`text-4xl font-black italic uppercase tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>Gerenciar Divisões</h2>
          <p className="text-slate-500 font-bold uppercase tracking-widest text-[10px] mt-2">Estrutura organizacional do dashboard</p>
        </div>
        <button 
          onClick={() => setIsModalOpen(true)}
          className="px-8 py-4 bg-brand-red text-white rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-2 hover:bg-red-600 transition-all active:scale-95 shadow-xl shadow-red-900/20"
        >
          <Plus className="w-5 h-5" />
          Nova Divisão
        </button>
      </header>

      <div className="grid gap-4">
        {data.map((section) => (
          <div key={section.id} className={`p-6 rounded-2xl border flex items-center justify-between group transition-all ${
            isWarRoom ? 'bg-slate-950/50 border-indigo-900/20 hover:border-indigo-500/40' : 'bg-slate-50/50 border-slate-100 hover:border-slate-300'
          }`}>
            {editingId === section.id ? (
              <input 
                autoFocus
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                onBlur={() => syncSectionName(section.id)}
                onKeyDown={(e) => e.key === 'Enter' && syncSectionName(section.id)}
                className={`bg-transparent font-black italic uppercase text-xl outline-none border-b-2 border-brand-red ${isWarRoom ? 'text-white' : 'text-slate-900'}`}
              />
            ) : (
              <div className="flex items-center gap-4">
                <div className="w-10 h-10 rounded-xl bg-brand-red/10 flex items-center justify-center text-brand-red font-black">
                  {section.metrics.length}
                </div>
                <h3 className={`text-xl font-black italic uppercase tracking-tight ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>{section.title}</h3>
              </div>
            )}
            
            <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
              <button 
                onClick={() => { setEditingId(section.id); setNewName(section.title); }}
                className={`p-3 rounded-xl transition-colors ${isWarRoom ? 'hover:bg-white/5 text-indigo-400' : 'hover:bg-white text-slate-400 hover:text-slate-900'}`}
              >
                <Edit2 className="w-5 h-5" />
              </button>
              <button 
                onClick={() => removeSection(section)}
                className={`p-3 rounded-xl transition-colors ${isWarRoom ? 'hover:bg-red-500/10 text-red-400' : 'hover:bg-white text-slate-400 hover:text-brand-red'}`}
              >
                <Trash2 className="w-5 h-5" />
              </button>
            </div>
          </div>
        ))}
      </div>

      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className={`w-full max-w-lg rounded-3xl p-8 ${isWarRoom ? 'bg-[#0f1125] text-white' : 'bg-white text-slate-900'}`}
            >
              <h3 className="text-2xl font-black italic uppercase mb-6">Nova Divisão</h3>
              <form onSubmit={saveSectionLocal} className="space-y-6">
                <input 
                  required
                  placeholder="Nome da Divisão"
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  className={`w-full px-6 py-4 rounded-2xl font-bold outline-none ${isWarRoom ? 'bg-slate-950 border-indigo-900/50' : 'bg-slate-50 border-slate-200'}`}
                />
                <div className="flex gap-4">
                  <button type="button" onClick={() => setIsModalOpen(false)} className="flex-1 py-4 border rounded-2xl font-black uppercase text-xs">Cancelar</button>
                  <button type="submit" className="flex-1 py-4 bg-brand-red text-white rounded-2xl font-black uppercase text-xs">Salvar</button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function MetricManager({ data, setData, isWarRoom, updateMetric, deleteMetric }: { 
  data: Section[], 
  setData: (d: Section[] | ((prev: Section[]) => Section[])) => void, 
  isWarRoom: boolean,
  updateMetric: (sectionId: string, metric: Metric) => Promise<void>,
  deleteMetric: (metricId: string, spId: number, resultTable?: string) => Promise<void>
}) {
  const [selectedSection, setSelectedSection] = useState<string>(data[0]?.id || '');
  const [editingMetric, setEditingMetric] = useState<Metric | null>(null);
  const [editingMetricSectionId, setEditingMetricSectionId] = useState<string>('');
  const [isModalOpen, setIsModalOpen] = useState(false);

  useEffect(() => {
    if (!selectedSection && data.length > 0) {
      setSelectedSection(data[0].id);
    }
  }, [data, selectedSection]);

  const saveMetricLocal = async (e: FormEvent) => {
    e.preventDefault();
    if (!editingMetric || !editingMetricSectionId) return;

    const metricId = editingMetric.id || `m-${Date.now()}`;
    const metricToSave: Metric = {
      ...editingMetric,
      id: metricId,
      status: editingMetric.status || 'info',
      lastUpdate: editingMetric.lastUpdate || new Date().toLocaleString('pt-BR'),
      history: editingMetric.history || []
    };

    await updateMetric(editingMetricSectionId, metricToSave);
    setIsModalOpen(false);
    setEditingMetric(null);
  };

  const removeMetric = async (metric: Metric) => {
    if (confirm('Excluir este card?')) {
      await deleteMetric(metric.id, metric.spId!, metric.resultTable);
    }
  };

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
       <header className="flex justify-between items-end mb-10">
        <div>
          <h2 className={`text-4xl font-black italic uppercase tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>Gerenciar Cards</h2>
          <p className="text-slate-500 font-bold uppercase tracking-widest text-[10px] mt-2">Métricas e regras de negócio</p>
        </div>
        <button 
          onClick={() => { 
            setEditingMetric({} as Metric); 
            setEditingMetricSectionId(selectedSection);
            setIsModalOpen(true); 
          }}
          className="px-8 py-4 bg-brand-red text-white rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-2 hover:bg-red-600 transition-all active:scale-95 shadow-xl shadow-red-900/20"
        >
          <Plus className="w-5 h-5" />
          Novo Card
        </button>
      </header>

      <div className="flex gap-4 mb-8 overflow-x-auto pb-2 scrollbar-hide">
        {data.map((s) => (
          <button 
            key={s.id}
            onClick={() => setSelectedSection(s.id)}
            className={`px-6 py-3 rounded-xl font-black text-[10px] uppercase tracking-[0.2em] whitespace-nowrap transition-all ${
              selectedSection === s.id 
              ? 'bg-brand-red text-white shadow-lg shadow-red-900/40' 
              : isWarRoom ? 'bg-white/5 text-slate-400 hover:bg-white/10' : 'bg-slate-100 text-slate-500 hover:bg-white border-slate-200'
            }`}
          >
            {s.title}
          </button>
        ))}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {data.find(s => s.id === selectedSection)?.metrics.map((metric) => (
          <div key={metric.id} className={`p-6 rounded-3xl border transition-all h-full flex flex-col justify-between group ${
            isWarRoom ? 'bg-[#0f1125] border-indigo-900/30 hover:border-indigo-500/50 text-white' : 'bg-white border-slate-200 shadow-sm hover:border-slate-400'
          }`}>
            <div>
              <h3 className="text-xl font-black italic uppercase tracking-tight mb-4">{metric.title}</h3>
              <div className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Objetivo</div>
              <p className={`text-xs font-semibold leading-relaxed mb-4 ${isWarRoom ? 'text-slate-400' : 'text-slate-500'} line-clamp-3`}>{metric.objective || 'Nenhum objetivo definido'}</p>
              
              <div className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${isWarRoom ? 'text-red-400' : 'text-brand-red'}`}>Query SQL</div>
              <div className={`p-3 rounded-xl font-mono text-[9px] break-all h-16 overflow-y-auto mb-4 ${isWarRoom ? 'bg-slate-950 text-indigo-300' : 'bg-slate-50 text-slate-600'}`}>
                {metric.query || '-- Nenhuma query definida'}
              </div>
            </div>

            <div className="flex gap-2 justify-end pt-4 border-t border-white/5">
              <button 
                onClick={() => { 
                  setEditingMetric(metric); 
                  setEditingMetricSectionId(selectedSection);
                  setIsModalOpen(true); 
                }}
                className={`p-3 rounded-xl transition-colors ${isWarRoom ? 'hover:bg-white/5 text-indigo-400' : 'hover:bg-slate-50 text-slate-400 hover:text-slate-900'}`}
              >
                <Edit2 className="w-5 h-5" />
              </button>
              <button 
                onClick={() => removeMetric(metric)}
                className={`p-3 rounded-xl transition-colors ${isWarRoom ? 'hover:bg-red-500/10 text-red-400' : 'hover:bg-slate-50 text-slate-400 hover:text-brand-red'}`}
              >
                <Trash2 className="w-5 h-5" />
              </button>
            </div>
          </div>
        ))}
      </div>

      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className={`w-full max-w-2xl max-h-[90vh] rounded-3xl shadow-2xl p-8 flex flex-col ${isWarRoom ? 'bg-[#0f1125] text-white border border-indigo-900/50' : 'bg-white text-slate-900'}`}
            >
              <div className="flex justify-between items-center mb-6 flex-shrink-0">
                <h3 className="text-2xl font-black italic uppercase tracking-tight">Configurar Card</h3>
                <button onClick={() => setIsModalOpen(false)}><X className="w-6 h-6" /></button>
              </div>

              <form onSubmit={saveMetricLocal} className="flex flex-col flex-grow overflow-hidden">
                <div className="flex-grow overflow-y-auto pr-2 space-y-6 scrollbar-hide py-2">
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Divisão Responsável</label>
                    <select 
                      required
                      value={editingMetricSectionId}
                      onChange={(e) => setEditingMetricSectionId(e.target.value)}
                      className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                    >
                      <option value="">Selecione a divisão</option>
                      {data.map(s => <option key={s.id} value={s.id}>{s.title}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Título do Card</label>
                    <input 
                      required
                      value={editingMetric?.title || ''}
                      onChange={(e) => setEditingMetric(prev => ({ ...prev!, title: e.target.value }))}
                      className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Objetivo da Métrica</label>
                    <textarea 
                      rows={2}
                      value={editingMetric?.objective || ''}
                      onChange={(e) => setEditingMetric(prev => ({ ...prev!, objective: e.target.value }))}
                      className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Regras Aplicadas (uma por linha)</label>
                    <textarea 
                      rows={2}
                      value={editingMetric?.rules?.join('\n') || ''}
                      onChange={(e) => setEditingMetric(prev => ({ ...prev!, rules: e.target.value.split('\n') }))}
                      className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Script de Querie (SQL)</label>
                    <textarea 
                      rows={3}
                      value={editingMetric?.query || ''}
                      onChange={(e) => setEditingMetric(prev => ({ ...prev!, query: e.target.value }))}
                      className={`w-full px-6 py-4 rounded-xl font-mono text-sm transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                    />
                  </div>
                  
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Intervalo de Atualização</label>
                    <select 
                      value={editingMetric?.refreshInterval || ''}
                      onChange={(e) => setEditingMetric(prev => ({ ...prev!, refreshInterval: parseInt(e.target.value) || undefined }))}
                      className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                    >
                      <option value="">Padrão (30 Segundos)</option>
                      <option value="300">5 Minutos</option>
                      <option value="900">15 Minutos</option>
                      <option value="1800">30 Minutos</option>
                      <option value="3600">1 Hora</option>
                    </select>
                  </div>
                </div>

                <div className="flex gap-4 pt-4 border-t border-white/5 flex-shrink-0 mt-4">
                  <button type="button" onClick={() => setIsModalOpen(false)} className={`flex-1 py-4 rounded-2xl font-black uppercase text-xs tracking-widest border transition-all ${isWarRoom ? 'border-slate-800 hover:bg-white/5' : 'border-slate-200 hover:bg-slate-50'}`}>Cancelar</button>
                  <button type="submit" className="flex-1 py-4 bg-brand-red text-white rounded-2xl font-black uppercase text-xs tracking-widest hover:bg-red-600 shadow-xl shadow-red-900/30 transition-all flex items-center justify-center gap-2">
                    <Save className="w-5 h-5" /> Salvar Card
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function UserManager({ users, setUsers, isWarRoom }: { users: User[], setUsers: (u: User[]) => void, isWarRoom: boolean }) {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<User | null>(null);

  const saveUser = (e: FormEvent) => {
    e.preventDefault();
    if (!editingUser) return;
    
    if (users.find(u => u.id === editingUser.id)) {
      setUsers(users.map(u => u.id === editingUser.id ? editingUser : u));
    } else {
      setUsers([...users, { ...editingUser, id: `u-${Date.now()}`, lastActive: 'Nunca', status: 'Ativo' }]);
    }
    setIsModalOpen(false);
    setEditingUser(null);
  };

  const deleteUser = (id: string) => {
    setUsers(users.filter(u => u.id !== id));
  };

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500 font-sans">
      <header className="flex justify-between items-end mb-10">
        <div>
          <h2 className={`text-4xl font-black italic uppercase tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>Gestão de Usuários</h2>
          <p className="text-slate-500 font-bold uppercase tracking-widest text-[10px] mt-2">Controle de acesso e permissões</p>
        </div>
        <button 
          onClick={() => { setEditingUser({ role: 'viewer' } as User); setIsModalOpen(true); }}
          className="px-8 py-4 bg-brand-red text-white rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-2 hover:bg-red-600 transition-all active:scale-95 shadow-xl shadow-red-900/20"
        >
          <Plus className="w-5 h-5" />
          Novo Usuário
        </button>
      </header>

      <div className={`rounded-3xl border overflow-hidden ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-200 shadow-sm'}`}>
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className={isWarRoom ? 'bg-slate-950/50' : 'bg-slate-50'}>
              <th className="px-8 py-6 text-[10px] font-black uppercase tracking-widest text-slate-500">Usuário</th>
              <th className="px-8 py-6 text-[10px] font-black uppercase tracking-widest text-slate-500">Email</th>
              <th className="px-8 py-6 text-[10px] font-black uppercase tracking-widest text-slate-500">Cargo</th>
              <th className="px-8 py-6 text-[10px] font-black uppercase tracking-widest text-slate-500">Último Acesso</th>
              <th className="px-8 py-6 text-[10px] font-black uppercase tracking-widest text-slate-500 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className={`divide-y ${isWarRoom ? 'divide-indigo-400/5' : 'divide-slate-100'}`}>
            {users.map((user) => (
              <tr key={user.id} className={`group transition-all ${isWarRoom ? 'hover:bg-white/[0.02]' : 'hover:bg-slate-50'}`}>
                <td className="px-8 py-6">
                  <div className="flex items-center gap-4">
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center font-black text-sm ${isWarRoom ? 'bg-indigo-500/10 text-indigo-400' : 'bg-slate-100 text-slate-600'}`}>
                      {user.name[0]}
                    </div>
                    <span className={`font-black italic uppercase tracking-tight text-lg ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>{user.name}</span>
                  </div>
                </td>
                <td className={`px-8 py-6 font-bold text-sm ${isWarRoom ? 'text-slate-400' : 'text-slate-500'}`}>{user.email}</td>
                <td className="px-8 py-6">
                  <span className={`px-4 py-1.5 rounded-full text-[9px] font-black uppercase tracking-[2px] border ${
                    user.role === 'admin' ? 'border-brand-red text-brand-red bg-brand-red/5' : 
                    user.role === 'editor' ? 'border-indigo-500 text-indigo-500 bg-indigo-500/5' : 
                    'border-slate-500 text-slate-500 bg-slate-500/5'
                  }`}>
                    {user.role}
                  </span>
                </td>
                <td className={`px-8 py-6 font-mono text-[11px] ${isWarRoom ? 'text-slate-500' : 'text-slate-400'}`}>{user.lastActive}</td>
                <td className="px-8 py-6 text-right">
                  <div className="flex gap-2 justify-end opacity-0 group-hover:opacity-100 transition-opacity">
                    <button onClick={() => { setEditingUser(user); setIsModalOpen(true); }} className={`p-2 rounded-lg transition-colors ${isWarRoom ? 'hover:bg-white/5 text-indigo-400' : 'hover:bg-white text-slate-400 hover:text-slate-900'}`}><Edit2 className="w-4 h-4" /></button>
                    <button onClick={() => deleteUser(user.id)} className={`p-2 rounded-lg transition-colors ${isWarRoom ? 'hover:bg-red-500/10 text-red-400' : 'hover:bg-white text-slate-400 hover:text-brand-red'}`}><Trash2 className="w-4 h-4" /></button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className={`w-full max-w-md rounded-3xl shadow-2xl p-8 flex flex-col ${isWarRoom ? 'bg-[#0f1125] text-white border border-indigo-900/50' : 'bg-white text-slate-900'}`}
            >
              <h3 className="text-2xl font-black italic uppercase tracking-tight mb-8">Conta de Usuário</h3>
              <form onSubmit={saveUser} className="space-y-6">
                <div>
                  <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Nome Completo</label>
                  <input 
                    required
                    value={editingUser?.name || ''}
                    onChange={(e) => setEditingUser(prev => ({ ...prev!, name: e.target.value }))}
                    className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Endereço de Email</label>
                  <input 
                    required
                    type="email"
                    value={editingUser?.email || ''}
                    onChange={(e) => setEditingUser(prev => ({ ...prev!, email: e.target.value }))}
                    className={`w-full px-6 py-4 rounded-2xl font-bold transition-all outline-none focus:ring-2 focus:ring-brand-red/20 ${isWarRoom ? 'bg-slate-950 border-indigo-900/50 text-white' : 'bg-slate-50 border-slate-200'}`}
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-2 block">Nível de Acesso</label>
                  <div className="grid grid-cols-3 gap-2">
                    {(['viewer', 'editor', 'admin'] as AccessLevel[]).map((role) => (
                      <button
                        key={role}
                        type="button"
                        onClick={() => setEditingUser(prev => ({ ...prev!, role }))}
                        className={`py-3 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all ${
                          editingUser?.role === role 
                          ? 'bg-brand-red text-white shadow-lg' 
                          : `${isWarRoom ? 'bg-slate-950 text-slate-500 border border-indigo-900/50' : 'bg-slate-50 text-slate-400'}`
                        }`}
                      >
                        {role}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="flex gap-4 pt-4">
                  <button type="button" onClick={() => setIsModalOpen(false)} className={`flex-1 py-4 rounded-2xl font-black uppercase text-xs tracking-widest border transition-all ${isWarRoom ? 'border-slate-800 hover:bg-white/5' : 'border-slate-200 hover:bg-slate-50'}`}>Cancelar</button>
                  <button type="submit" className="flex-1 py-4 bg-brand-red text-white rounded-2xl font-black uppercase text-xs tracking-widest hover:bg-red-600 shadow-xl shadow-red-900/30 transition-all">Salvar</button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function DashboardView({ data, isWarRoom, refreshCountdown, isRefreshing, refreshData, totalDivergences, criticalMetrics, handleWidthChange, getPackedSections, setSelectedMetric, isAdmin, metricCountdowns }: { 
  data: Section[], 
  isWarRoom: boolean, 
  refreshCountdown: number, 
  isRefreshing: boolean, 
  refreshData: () => void, 
  totalDivergences: number, 
  criticalMetrics: number,
  handleWidthChange: (title: string, width: number) => void,
  getPackedSections: () => any[],
  setSelectedMetric: (m: Metric) => void,
  isAdmin: boolean,
  metricCountdowns: Record<string, number>
}) {
  return (
    <div className="space-y-10 animate-in fade-in duration-700">
      <header className={`flex flex-col sm:flex-row justify-between items-center mb-6 px-6 py-6 rounded-2xl transition-all duration-700 gap-6 mx-auto w-full ${
        isWarRoom 
        ? 'bg-[#0f1125] border border-indigo-900/40 shadow-[0_0_40px_rgba(0,0,0,0.5)] max-w-[1700px]' 
        : 'bg-white shadow-sm border border-slate-200 max-w-[1400px]'
      }`}>
        <div className="flex items-center gap-4">
          <div className={`p-3 rounded-2xl transition-colors duration-500 relative ${isWarRoom ? 'bg-indigo-950/30' : 'bg-slate-50'}`}>
            <Activity className={`w-8 h-8 transition-colors duration-500 ${isWarRoom ? 'text-brand-red animate-pulse' : 'text-slate-900'}`} />
            {isWarRoom && (
              <motion.div 
                animate={{ scale: [1, 1.5, 1], opacity: [0.5, 0, 0.5] }}
                transition={{ repeat: Infinity, duration: 2 }}
                className="absolute inset-0 bg-brand-red/20 rounded-2xl"
              />
            )}
          </div>
          <div>
            <h1 className={`text-2xl font-black italic tracking-tighter uppercase leading-tight transition-colors duration-500 ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>
              Monitor <span className="text-brand-red font-black">Operational</span>
            </h1>
            <div className="flex items-center gap-2 mt-1">
              <div className={`w-2 h-2 rounded-full animate-pulse ${isWarRoom ? 'bg-red-500 shadow-[0_0_8px_rgba(239,68,68,0.8)]' : 'bg-emerald-500'}`} />
              <p className={`text-[10px] font-black uppercase tracking-[0.2em] ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Command Center &bull; Live Stream</p>
            </div>
          </div>
        </div>

        {/* Sync Countdown Indicator */}
        <div className="hidden md:flex flex-col items-center gap-1 group">
          <div className="flex items-center gap-2">
            <Clock className={`w-3 h-3 ${isWarRoom ? 'text-indigo-500' : 'text-slate-300'}`} />
            <span className={`text-[10px] font-black tabular-nums transition-colors ${isWarRoom ? 'text-white' : 'text-slate-500'}`}>
              Próximo Sync: {refreshCountdown}s
            </span>
          </div>
          <div className={`w-32 h-1 rounded-full overflow-hidden transition-colors ${isWarRoom ? 'bg-slate-800' : 'bg-slate-100'}`}>
            <motion.div 
              initial={{ width: '100%' }}
              animate={{ width: `${(refreshCountdown / 30) * 100}%` }}
              transition={{ ease: "linear" }}
              className={`h-full ${isWarRoom ? 'bg-brand-red shadow-[0_0_8px_rgba(239,68,68,0.5)]' : 'bg-brand-red'}`} 
            />
          </div>
        </div>
        
        <div className="flex items-center gap-3">
          <button 
            onClick={() => {}} // This should be handled in App
            className={`flex items-center gap-2 px-6 py-3 rounded-xl font-black text-xs uppercase tracking-widest transition-all active:scale-95 ${
              isWarRoom 
              ? 'bg-brand-red text-white hover:bg-red-600 shadow-[0_0_20px_rgba(204,0,0,0.4)]' 
              : 'bg-slate-900 text-white hover:bg-slate-800'
            }`}
             id="war-room-toggle"
          >
            {isWarRoom ? 'Sair do Modo War Room' : 'Ativar Modo War Room'}
          </button>

          <button 
            onClick={refreshData}
            disabled={isRefreshing}
            className={`flex items-center gap-2 px-4 py-3 rounded-xl transition-all text-xs font-black border ${
              isWarRoom 
              ? 'bg-indigo-950/20 border-indigo-900/50 text-indigo-300 hover:text-white hover:bg-indigo-900/30' 
              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
            } disabled:opacity-50`}
          >
            <RefreshCcw className={`w-4 h-4 ${isRefreshing ? 'animate-spin' : ''}`} />
            {isRefreshing ? 'SINCRONIZANDO...' : 'SINCRONIZAR'}
          </button>
        </div>
      </header>

      {/* Global Performance KPIs */}
      <div className={`grid grid-cols-2 md:grid-cols-4 gap-4 mx-auto w-full transition-all duration-700 ${isWarRoom ? 'max-w-[1700px]' : 'max-w-[1400px]'}`}>
        <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-100 shadow-sm'}`}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Total Divergências</p>
              <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>{totalDivergences}</h4>
            </div>
            <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-indigo-500/10 text-indigo-400' : 'bg-slate-50 text-slate-400'}`}>
              <Activity className="w-5 h-5" />
            </div>
          </div>
          <div className={`mt-4 h-1.5 w-full rounded-full overflow-hidden ${isWarRoom ? 'bg-slate-800' : 'bg-slate-100'}`}>
            <motion.div initial={{ width: 0 }} animate={{ width: '65%' }} className="h-full bg-brand-red" />
          </div>
        </div>

        <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-100 shadow-sm'}`}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Métricas Críticas</p>
              <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-brand-red' : 'text-brand-red'}`}>{criticalMetrics}</h4>
            </div>
            <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-red-500/10 text-red-400' : 'bg-red-50 text-red-400'}`}>
              <XCircle className="w-5 h-5" />
            </div>
          </div>
          <p className="mt-4 text-[10px] font-bold text-slate-500 italic">Requer atenção imediata</p>
        </div>

        <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-100 shadow-sm'}`}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Acertos Hoje</p>
              <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-emerald-500' : 'text-emerald-600'}`}>88%</h4>
            </div>
            <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-emerald-500/10 text-emerald-400' : 'bg-emerald-50 text-emerald-600'}`}>
              <CheckCircle2 className="w-5 h-5" />
            </div>
          </div>
          <div className="mt-4 flex items-center gap-1">
            <TrendingUp className="w-3 h-3 text-emerald-500" />
            <span className="text-[10px] font-black text-emerald-500">+12.5% vs ontem</span>
          </div>
        </div>

        <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-brand-red border-red-800 shadow-[0_0_30px_rgba(204,0,0,0.2)]' : 'bg-brand-yellow border-brand-yellow shadow-sm'}`}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>SLA Médio</p>
              <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>18m</h4>
            </div>
            <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-white/10 text-white' : 'bg-white/20 text-slate-900'}`}>
              <RefreshCcw className="w-5 h-5" />
            </div>
          </div>
          <p className={`mt-4 text-[10px] font-black uppercase ${isWarRoom ? 'text-red-200' : 'text-slate-700'}`}>Status: Excelente</p>
        </div>
      </div>

      <div className={`flex flex-wrap items-stretch gap-x-12 gap-y-12 px-2 transition-all duration-700 mx-auto ${isWarRoom ? 'max-w-[1700px]' : 'max-w-[1400px]'}`}>
        <AnimatePresence mode="popLayout">
          {getPackedSections().map(({ section, config }) => (
            <motion.div 
              layout
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              key={section.title} 
              className="transition-all duration-500 flex flex-col" 
              style={{ 
                width: `calc(${config.width}% - ${config.width === 100 ? '0px' : '24px'})`,
                minWidth: '320px',
                flexGrow: 1
              }}
            >
              <SectionContainer 
                section={section} 
                onCardClick={setSelectedMetric} 
                isWarRoom={isWarRoom} 
                metricCountdowns={metricCountdowns}
              />
            </motion.div>
          ))}
        </AnimatePresence>
      </div>
    </div>
  );
}

export default function App() {
  const navigate = useNavigate();
  const location = useLocation();
  const { 
    data, 
    setData, 
    users, 
    setUsers, 
    loading, 
    updateMetric, 
    deleteMetric, 
    updateSection, 
    deleteSection, 
    executeMetricQuery 
  } = useDataPersistence();
  
  const [currentView, setCurrentView] = useState('dashboard');
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [selectedMetric, setSelectedMetric] = useState<Metric | null>(null);
  const [isWarRoom, setIsWarRoom] = useState(true);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [settingsTab, setSettingsTab] = useState<'layout' | 'log'>('layout');
  const [layoutConfig, setLayoutConfig] = useState<{title: string, width: number}[]>([]);
  const [refreshCountdown, setRefreshCountdown] = useState(30);
  const [metricCountdowns, setMetricCountdowns] = useState<Record<string, number>>({});
  const [eventLog, setEventLog] = useState<{ id: string, message: string, time: string, type: 'info' | 'critical' | 'success' }[]>([]);

  useEffect(() => {
    if (data.length > 0 && layoutConfig.length === 0) {
      setLayoutConfig(data.map(s => ({ title: s.title, width: s.title === 'Validação de saldo' ? 100 : 50 })));
    }
  }, [data, layoutConfig.length]);

  // Set view based on URL
  useEffect(() => {
    const path = location.pathname.substring(1);
    if (path === 'dashboard' || path === 'live') {
      setCurrentView('dashboard');
    } else if (['divisions', 'cards', 'users'].includes(path)) {
      setCurrentView(path as any);
    } else if (path === '') {
      setCurrentView('dashboard');
    }
  }, [location.pathname]);

  const setView = (view: any) => {
    setCurrentView(view);
    navigate(`/${view === 'dashboard' ? '' : view}`);
  };

  // Individual metric refresh logic
  useEffect(() => {
    const timer = setInterval(() => {
      setMetricCountdowns(prev => {
        const next = { ...prev };
        data.forEach(s => {
          s.metrics.forEach(m => {
            if (m.refreshInterval) {
              if (next[m.id] === undefined) {
                next[m.id] = m.refreshInterval;
              } else if (next[m.id] <= 1) {
                // Refresh execution
                executeMetricQuery(m).then(result => {
                  if (result) {
                    setData(prevData => prevData.map(ps => ({
                      ...ps,
                      metrics: ps.metrics.map(pm => pm.id === m.id ? { 
                        ...pm, 
                        value: result.value, 
                        status: result.status, 
                        lastUpdate: new Date().toLocaleString('pt-BR'),
                        details: result.details 
                      } : pm)
                    })));
                  }
                });
                next[m.id] = m.refreshInterval;
              } else {
                next[m.id] -= 1;
              }
            } else if (next[m.id] !== undefined) {
              delete next[m.id];
            }
          });
        });
        return next;
      });
    }, 1000);
    return () => clearInterval(timer);
  }, [data, executeMetricQuery, setData]);

  // Global Auto-refresh logic
  useEffect(() => {
    const timer = setInterval(() => {
      setRefreshCountdown((prev) => {
        if (prev <= 1) {
          refreshData();
          return 30;
        }
        return prev - 1;
      });
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  // Subtle data fluctuation (kept from original)
  useEffect(() => {
    const flucInterval = setInterval(() => {
      setData(prev => {
        const newEvents: typeof eventLog = [];
        const nextState = prev.map(section => ({
          ...section,
          metrics: section.metrics.map(m => {
            if (Math.random() > 0.9) {
              const delta = Math.floor(Math.random() * 5) - 2; 
              const newValue = Math.max(0, m.value + delta);
              
              if (newValue !== m.value) {
                const type = newValue === 0 ? 'success' : (newValue > 15 ? 'critical' : 'info');
                newEvents.push({
                  id: Math.random().toString(36).substr(2, 9),
                  message: `${m.title}: ${m.value} -> ${newValue}`,
                  time: new Date().toLocaleTimeString('pt-BR'),
                  type
                });
              }

              const newHistoryPoint: MetricHistory = {
                date: new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
                value: newValue
              };

              return {
                ...m,
                value: newValue,
                history: [...(m.history || []).slice(1), newHistoryPoint],
                status: newValue > 15 ? 'critical' : (newValue > 0 ? 'info' : 'ok')
              };
            }
            return m;
          })
        }));

        if (newEvents.length > 0) {
          setEventLog(prevEvents => [ ...newEvents, ...prevEvents].slice(0, 50));
        }

        return nextState;
      });
    }, 4000);
    return () => clearInterval(flucInterval);
  }, []);

  const refreshData = () => {
    setIsRefreshing(true);
    setTimeout(() => {
      setData(prev => prev.map(section => ({
        ...section,
        metrics: section.metrics.map(m => ({
          ...m,
          lastUpdate: new Date().toLocaleString('pt-BR')
        }))
      })));
      setIsRefreshing(false);
    }, 800);
  };

  const totalDivergences = data.reduce((acc, section) => acc + section.metrics.reduce((mAcc, m) => mAcc + Number(m.value) || 0, 0), 0);
  const criticalMetrics = data.reduce((acc, section) => acc + section.metrics.filter(m => m.status === 'critical').length, 0);

  const handleWidthChange = (title: string, width: number) => {
    setLayoutConfig(prev => prev.map(c => c.title === title ? { ...c, width } : c));
  };

  const getPackedSections = () => {
    const sectionsWithConfig = layoutConfig.map(c => ({
      config: c,
      section: data.find(s => s.title === c.title)!
    })).filter(item => item.section);

    const result: (typeof sectionsWithConfig[0])[] = [];
    const remaining = [...sectionsWithConfig];
    
    while (remaining.length > 0) {
      let currentRowWidth = 0;
      const rowIndices: number[] = [];
      for (let i = 0; i < remaining.length; i++) {
        if (currentRowWidth + remaining[i].config.width <= 100.1) {
          currentRowWidth += remaining[i].config.width;
          rowIndices.push(i);
        }
      }
      rowIndices.sort((a, b) => b - a).forEach(idx => {
        result.push(remaining[idx]);
        remaining.splice(idx, 1);
      });
      if (rowIndices.length === 0 && remaining.length > 0) {
        result.push(remaining.shift()!);
      }
    }
    return result;
  };

  const isLiveView = location.pathname === '/live';

  if (loading) {
    return (
      <div className={`min-h-screen flex flex-col items-center justify-center transition-colors duration-700 ${isWarRoom ? 'bg-[#050510] text-white' : 'bg-slate-50 text-slate-900'}`}>
        <RefreshCcw className="w-12 h-12 text-brand-red animate-spin mb-4" />
        <h2 className="text-xl font-bold tracking-widest uppercase italic">Carregando Configurações...</h2>
        <p className="text-slate-500 text-xs mt-2">Sincronizando com SharePoint</p>
      </div>
    );
  }

  return (
    <div className={`min-h-screen transition-colors duration-700 ${isWarRoom ? 'bg-[#050510] text-white' : 'bg-slate-50 text-slate-900'}`}>
      {!isLiveView && currentView !== 'dashboard' && (
        <Sidebar currentView={currentView} setView={setView} isWarRoom={isWarRoom} />
      )}

      <main className={`${(currentView === 'dashboard' || isLiveView) ? 'w-full p-4 md:p-8' : 'ml-20 md:ml-64 p-4 md:p-10'} pb-32 transition-all duration-700 overflow-hidden`}>
        <Routes>
          <Route path="/live" element={
            <DashboardView 
              data={data}
              isWarRoom={isWarRoom}
              refreshCountdown={refreshCountdown}
              isRefreshing={isRefreshing}
              refreshData={refreshData}
              totalDivergences={totalDivergences}
              criticalMetrics={criticalMetrics}
              handleWidthChange={handleWidthChange}
              getPackedSections={getPackedSections}
              setSelectedMetric={setSelectedMetric}
              isAdmin={false}
              metricCountdowns={metricCountdowns}
            />
          } />
          <Route path="/" element={
            <div className="relative">
              {/* Top Navigation to System */}
              <div className="flex justify-end mb-4 px-2">
                <button 
                  onClick={() => setView('divisions')}
                  className={`flex items-center gap-2 px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all ${
                    isWarRoom ? 'bg-white/5 text-slate-400 hover:text-white hover:bg-white/10' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                  }`}
                >
                  <Settings className="w-3 h-3" />
                  Painel Administrativo
                </button>
              </div>

              <DashboardView 
                data={data}
                isWarRoom={isWarRoom}
                refreshCountdown={refreshCountdown}
                isRefreshing={isRefreshing}
                refreshData={refreshData}
                totalDivergences={totalDivergences}
                criticalMetrics={criticalMetrics}
                handleWidthChange={handleWidthChange}
                getPackedSections={getPackedSections}
                setSelectedMetric={setSelectedMetric}
                isAdmin={true}
                metricCountdowns={metricCountdowns}
              />
            </div>
          } />
          <Route path="/divisions" element={<SectionManager data={data} setData={setData} isWarRoom={isWarRoom} updateSection={updateSection} deleteSection={deleteSection} />} />
          <Route path="/cards" element={<MetricManager data={data} setData={setData} isWarRoom={isWarRoom} updateMetric={updateMetric} deleteMetric={deleteMetric} />} />
          <Route path="/users" element={<UserManager users={users} setUsers={setUsers} isWarRoom={isWarRoom} />} />
        </Routes>

        {/* Global Modal for Metric Details (from Dashboard) */}
        <AnimatePresence>
          {selectedMetric && (
            <DivergenceModal 
              metric={selectedMetric} 
              onClose={() => setSelectedMetric(null)} 
            />
          )}
        </AnimatePresence>

        {/* Settings Panel (Moved toggle to Sidebar or floating if needed) */}
        <AnimatePresence>
          {isSettingsOpen && (
            <>
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setIsSettingsOpen(false)}
                className="fixed inset-0 bg-slate-950/20 backdrop-blur-[2px] z-[105]"
              />
              <motion.div
                initial={{ opacity: 0, x: 300 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 300 }}
                className={`fixed right-0 top-0 bottom-0 w-80 z-[110] shadow-2xl p-6 border-l transition-colors duration-500 flex flex-col ${isWarRoom ? 'bg-slate-900 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'}`}
              >
                <div className="flex justify-between items-center mb-6">
                  <div className="flex items-center gap-3">
                    <Activity className="w-6 h-6 text-brand-red" />
                    <h2 className="text-xl font-black italic uppercase tracking-tighter">Centro de Controle</h2>
                  </div>
                  <button onClick={() => setIsSettingsOpen(false)} className="p-2 rounded-full hover:bg-slate-100 text-slate-500">
                    <X className="w-6 h-6" />
                  </button>
                </div>
                {/* Tabs */}
                <div className={`flex p-1 rounded-xl mb-8 transition-colors ${isWarRoom ? 'bg-slate-950' : 'bg-slate-100'}`}>
                  <button 
                    onClick={() => setSettingsTab('layout')}
                    className={`flex-1 py-2 text-[10px] font-black uppercase tracking-widest rounded-lg transition-all ${settingsTab === 'layout' ? 'bg-brand-red text-white shadow-md' : 'text-slate-500 hover:bg-white/10'}`}
                  >
                    Layout
                  </button>
                  <button 
                    onClick={() => setSettingsTab('log')}
                    className={`flex-1 py-2 text-[10px] font-black uppercase tracking-widest rounded-lg transition-all ${settingsTab === 'log' ? 'bg-brand-red text-white shadow-md' : 'text-slate-500 hover:bg-white/10'}`}
                  >
                    Atividade
                  </button>
                </div>

                <div className="flex-grow overflow-y-auto pr-2 scrollbar-hide">
                  {settingsTab === 'layout' ? (
                    <div className="space-y-8 animate-in fade-in slide-in-from-right-4 duration-300">
                      {layoutConfig.map((config) => (
                        <div key={config.title} className="space-y-4">
                          <div className="flex justify-between items-end">
                            <p className={`text-[10px] font-black uppercase tracking-widest leading-relaxed max-w-[70%] ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>
                              {config.title}
                            </p>
                            <span className="text-lg font-black tabular-nums">{config.width}%</span>
                          </div>
                          <input 
                            type="range" 
                            min="30" max="100" step="5"
                            value={config.width}
                            onChange={(e) => handleWidthChange(config.title, parseInt(e.target.value))}
                            className={`w-full h-1.5 rounded-lg appearance-none cursor-pointer accent-brand-red ${isWarRoom ? 'bg-slate-800' : 'bg-slate-100'}`}
                          />
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="space-y-3 animate-in fade-in slide-in-from-left-4 duration-300">
                      {eventLog.length > 0 ? (
                        eventLog.map((event) => (
                          <div key={event.id} className={`p-3 rounded-xl border-l-4 transition-all ${
                            isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-200'
                          } ${
                            event.type === 'critical' ? 'border-l-red-500' : 
                            event.type === 'success' ? 'border-l-emerald-500' : 'border-l-indigo-500'
                          }`}>
                            <div className="flex justify-between items-center mb-1">
                              <span className={`text-[8px] font-black uppercase tracking-widest ${isWarRoom ? 'text-slate-500' : 'text-slate-400'}`}>{event.time}</span>
                              {event.type === 'critical' && <Triangle className="w-2 h-2 text-red-500 fill-red-500 animate-pulse" />}
                            </div>
                            <p className={`text-[10px] font-bold leading-tight ${isWarRoom ? 'text-slate-300' : 'text-slate-700'}`}>{event.message}</p>
                          </div>
                        ))
                      ) : (
                        <div className="text-center py-20">
                          <p className="text-slate-500 font-bold italic text-xs">Nenhuma atividade recente detectada.</p>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </motion.div>
            </>
          )}
        </AnimatePresence>

        <div className="fixed bottom-8 right-8 z-[90] flex flex-col gap-4">
          <button
            onClick={() => setIsWarRoom(!isWarRoom)}
            className={`p-4 rounded-2xl shadow-2xl flex items-center gap-3 transition-all active:scale-95 group overflow-hidden ${
              isWarRoom 
              ? 'bg-brand-red text-white hover:bg-red-600 shadow-[0_10px_30px_rgba(204,0,0,0.5)]' 
              : 'bg-slate-900 text-white hover:bg-slate-800 shadow-[0_10px_30px_rgba(0,0,0,0.3)]'
            }`}
          >
            <Activity className={`w-6 h-6 ${isWarRoom ? 'animate-pulse' : ''}`} />
            <span className="font-black text-xs uppercase tracking-widest pr-2 border-l border-white/20 pl-3">WAR ROOM</span>
          </button>
          
          <button
            onClick={() => setIsSettingsOpen(true)}
            className={`p-4 rounded-2xl shadow-2xl flex items-center gap-3 transition-all active:scale-95 group overflow-hidden ${
              isWarRoom 
              ? 'bg-indigo-950/80 text-white border border-indigo-900/40' 
              : 'bg-white text-slate-900 border border-slate-200'
            }`}
          >
            <Settings className="w-6 h-6 group-hover:rotate-180 transition-transform duration-700" />
            <span className="font-black text-xs uppercase tracking-widest pr-2 border-l border-slate-200 pl-3">Ajustar Dashboard</span>
          </button>
        </div>

        <footer className={`text-center text-[10px] py-10 border-t mt-12 transition-colors duration-500 mx-auto w-full ${
          isWarRoom ? 'border-indigo-950 text-indigo-900' : 'border-slate-200 text-slate-400'
        }`}>
          &copy; {new Date().getFullYear()} Monitoring System - Todos os direitos reservados
        </footer>
      </main>
    </div>
  );
}

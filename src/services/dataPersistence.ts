import { useState, useEffect, useCallback } from 'react';
import * as sp from './spService';
import { Section, Metric, User } from '../types';
import { postSqlQuery } from './apiService';

const LIST_DIVISIONS = 'DivisionsConfig';
const LIST_CARDS = 'CardsConfig';
const LIST_USERS = 'UsersConfig';

export function useDataPersistence() {
  const [data, setData] = useState<Section[]>([]);
  const [users, setUsers] = useState<User[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const ensureLists = useCallback(async () => {
    try {
      // Ensure Divisions List
      if (!(await sp.spListExists(LIST_DIVISIONS))) {
        await sp.spCreateList(LIST_DIVISIONS);
        await sp.spListEnsureTextField(LIST_DIVISIONS, 'Title'); // For Section Title
        await sp.spListEnsureTextField(LIST_DIVISIONS, 'SectionId');
      }

      // Ensure Cards List
      if (!(await sp.spListExists(LIST_CARDS))) {
        await sp.spCreateList(LIST_CARDS);
        await sp.spListEnsureTextField(LIST_CARDS, 'Title');
        await sp.spListEnsureTextField(LIST_CARDS, 'MetricId');
        await sp.spListEnsureTextField(LIST_CARDS, 'SectionId');
        await sp.spListEnsureMultiLineTextField(LIST_CARDS, 'Objective');
        await sp.spListEnsureMultiLineTextField(LIST_CARDS, 'Rules');
        await sp.spListEnsureMultiLineTextField(LIST_CARDS, 'SqlQuery');
        await sp.spListEnsureNumberField(LIST_CARDS, 'RefreshInterval');
        await sp.spListEnsureTextField(LIST_CARDS, 'ResultTable');
      }

      // Ensure Users List
      if (!(await sp.spListExists(LIST_USERS))) {
        await sp.spCreateList(LIST_USERS);
        await sp.spListEnsureTextField(LIST_USERS, 'Title'); // Name
        await sp.spListEnsureTextField(LIST_USERS, 'Email');
        await sp.spListEnsureTextField(LIST_USERS, 'Role');
        await sp.spListEnsureTextField(LIST_USERS, 'Status');
      }
    } catch (err: any) {
      console.error('Error ensuring lists:', err);
    }
  }, []);

  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      await ensureLists();

      const [divRes, cardRes, userRes] = await Promise.all([
        sp.spListGetItems(LIST_DIVISIONS),
        sp.spListGetItems(LIST_CARDS),
        sp.spListGetItems(LIST_USERS)
      ]);

      if (divRes.status && cardRes.status) {
        const sections: Section[] = divRes.data.map((d: any) => ({
          id: d.SectionId,
          spId: d.Id,
          title: d.Title,
          metrics: []
        }));

        cardRes.data.forEach((c: any) => {
          const section = sections.find(s => s.id === c.SectionId);
          if (section) {
            section.metrics.push({
              id: c.MetricId,
              spId: c.Id,
              title: c.Title,
              value: '---', // Initial value
              status: 'info',
              lastUpdate: 'Nunca',
              objective: c.Objective,
              rules: c.Rules ? c.Rules.split('\n') : [],
              query: c.SqlQuery,
              refreshInterval: c.RefreshInterval,
              resultTable: c.ResultTable
            });
          }
        });

        setData(sections);
      }

      if (userRes.status) {
        setUsers(userRes.data.map((u: any) => ({
          id: String(u.Id),
          spId: u.Id,
          name: u.Title,
          email: u.Email,
          role: u.Role,
          status: u.Status,
          lastActive: 'Desconhecido'
        })));
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, [ensureLists]);

  useEffect(() => {
    if (sp.hasSpContext()) {
      loadData();
    } else {
      // Fallback for local development if needed, or error
      setError('Ambiente SharePoint não detectado.');
      setLoading(false);
    }
  }, [loadData]);

  const updateMetric = async (sectionId: string, metric: Metric) => {
    try {
      const isNew = !metric.spId;
      const resultListName = `CardData_${metric.id.replace(/-/g, '_')}`;

      if (isNew) {
        // Create result list
        if (!(await sp.spListExists(resultListName))) {
          await sp.spCreateList(resultListName);
          await sp.spListEnsureTextField(resultListName, 'Title'); // Field name
          await sp.spListEnsureTextField(resultListName, 'Value'); // Field value
          await sp.spListEnsureTextField(resultListName, 'Status'); 
          await sp.spListEnsureMultiLineTextField(resultListName, 'RawData'); // JSON raw
        }

        const addRes = await sp.spListAddItem(LIST_CARDS, {
          Title: metric.title,
          MetricId: metric.id,
          SectionId: sectionId,
          Objective: metric.objective || '',
          Rules: metric.rules?.join('\n') || '',
          SqlQuery: metric.query || '',
          RefreshInterval: metric.refreshInterval || 0,
          ResultTable: resultListName
        });
        if (addRes.status) {
          metric.spId = addRes.data.id;
          metric.resultTable = resultListName;
        }
      } else {
        await sp.spListUpdateItem(LIST_CARDS, metric.spId!, {
          Title: metric.title,
          Objective: metric.objective || '',
          Rules: metric.rules?.join('\n') || '',
          SqlQuery: metric.query || '',
          RefreshInterval: metric.refreshInterval || 0
        });
      }
      
      // Update local state is handled in App.tsx by calling refresh/reload or manual state update
      await loadData(); 
    } catch (err) {
      console.error('Failed to update metric in SP:', err);
    }
  };

  const deleteMetric = async (metricId: string, spId: number, resultTable?: string) => {
    try {
      await sp.spListDeleteItem(LIST_CARDS, spId);
      if (resultTable) {
        await sp.spDeleteList(resultTable);
      }
      await loadData();
    } catch (err) {
      console.error('Failed to delete metric from SP:', err);
    }
  };

  const updateSection = async (section: Section) => {
     try {
       if (!section.spId) {
         await sp.spListAddItem(LIST_DIVISIONS, {
           Title: section.title,
           SectionId: section.id
         });
       } else {
         await sp.spListUpdateItem(LIST_DIVISIONS, section.spId, {
           Title: section.title
         });
       }
       await loadData();
     } catch (err) {
       console.error('Failed to update section in SP:', err);
     }
  };

  const deleteSection = async (section: Section) => {
    try {
      if (section.spId) {
        // Also delete all metrics in this section
        for (const m of section.metrics) {
          if (m.spId) await sp.spListDeleteItem(LIST_CARDS, m.spId);
          if (m.resultTable) await sp.spDeleteList(m.resultTable);
        }
        await sp.spListDeleteItem(LIST_DIVISIONS, section.spId);
      }
      await loadData();
    } catch (err) {
       console.error('Failed to delete section in SP:', err);
    }
  };

  const executeMetricQuery = async (metric: Metric) => {
    if (!metric.query) return;
    try {
      const data = await postSqlQuery(metric.query, metric.id);
      
      // Assume data[0] or sum logic for the card value
      let value: string | number = '---';
      let status: 'ok' | 'critical' | 'info' = 'info';

      if (Array.isArray(data) && data.length > 0) {
        // Heuristic: if there is a 'VALUE' or 'VALOR' field, use it. Otherwise use object count.
        const item = data[0];
        const valField = Object.keys(item).find(k => k.toUpperCase() === 'VALUE' || k.toUpperCase() === 'VALOR');
        value = valField ? item[valField] : data.length;
        
        // Status heuristic: if value is number and > 100 ? critical (just for demo/default)
        if (typeof value === 'number' && value > 100) status = 'critical';
        else if (data.length > 0) status = 'ok';
      }

      // Save to card result list
      if (metric.resultTable) {
        // Clear old results or just add new? 
        // Instructions: "puxa os resultados dessa lista até a proxima atualização"
        // Usually, clear and add new or just add with a timestamp and query latest.
        // Let's just add the latest result record.
        await sp.spListAddItem(metric.resultTable, {
          Title: 'LastSync',
          Value: String(value),
          Status: status,
          RawData: JSON.stringify(data)
        });
      }

      return { value, status, details: data };
    } catch (err) {
      console.error(`Error executing query for ${metric.title}:`, err);
      return null;
    }
  };

  return {
    data,
    users,
    loading,
    error,
    updateMetric,
    deleteMetric,
    updateSection,
    deleteSection,
    executeMetricQuery,
    loadData,
    setData,
    setUsers
  };
}

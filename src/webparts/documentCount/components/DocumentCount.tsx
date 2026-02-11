import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './DocumentCount.module.scss';

export interface IDocumentCountProps {
  libraryTitle: string;
  refreshIntervalSeconds: number;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

const MINIMUM_REFRESH_SECONDS = 10;

const DocumentCount: React.FC<IDocumentCountProps> = ({
  libraryTitle,
  refreshIntervalSeconds,
  spHttpClient,
  siteUrl
}) => {
  const [count, setCount] = React.useState<number | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    let isActive = true;

    const refresh = async () => {
      if (!libraryTitle) {
        if (isActive) {
          setError('Provide a library title to load the document count.');
          setIsLoading(false);
        }
        return;
      }

      if (isActive) {
        setIsLoading(true);
      }

      try {
        const safeTitle = encodeURIComponent(libraryTitle.replace(/'/g, "''"));
        const response = await spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${safeTitle}')?$select=ItemCount`,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          throw new Error(`Request failed with status ${response.status}`);
        }

        const payload: { ItemCount?: number } = await response.json();
        if (isActive) {
          setCount(payload.ItemCount ?? 0);
          setError(null);
        }
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Unknown error';
        if (isActive) {
          setError(`Unable to load documents: ${message}`);
        }
      } finally {
        if (isActive) {
          setIsLoading(false);
        }
      }
    };

    void refresh();

    const intervalSeconds = Math.max(refreshIntervalSeconds || 0, MINIMUM_REFRESH_SECONDS);
    const intervalId = window.setInterval(refresh, intervalSeconds * 1000);

    return () => {
      isActive = false;
      window.clearInterval(intervalId);
    };
  }, [libraryTitle, refreshIntervalSeconds, siteUrl, spHttpClient]);

  return (
    <section className={styles.documentCount}>
      <header className={styles.header}>
        <h2 className={styles.title}>Library document count</h2>
        <p className={styles.subtitle}>
          {libraryTitle ? `Tracking: ${libraryTitle}` : 'Set a library title in the web part settings.'}
        </p>
      </header>
      <div className={styles.countCard}>
        {isLoading ? (
          <span className={styles.loading}>Loading…</span>
        ) : (
          <span className={styles.countValue}>{count ?? 0}</span>
        )}
        <span className={styles.countLabel}>Documents</span>
      </div>
      {error && <p className={styles.error}>{error}</p>}
      <p className={styles.helper}>Updates automatically every few seconds without refreshing the page.</p>
    </section>
  );
};

export default DocumentCount;

import { Component, ReactNode } from 'react';
import { getStore } from '../store';

interface State {
  error: Error | null;
}

/**
 * Catches rendering errors so a bug never leaves a blank page. The workbook lives outside React,
 * so it is still intact: the fallback offers a backup download before reloading.
 */
export class ErrorBoundary extends Component<{ children: ReactNode }, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error) {
    console.error('GRID hit an unexpected error', error);
  }

  private backup = () => {
    try {
      const wb = getStore().wb;
      const blob = new Blob([JSON.stringify(wb.serialize())], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `${(wb.title || 'workbook').replace(/[\\/:*?"<>|]+/g, '_')} (backup).grid`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    } catch (e) {
      console.error('Backup failed', e);
    }
  };

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="crash" role="alert">
        <div className="crash-card">
          <h1>Something went wrong</h1>
          <p>GRID ran into an unexpected problem. Your data is still in memory: download a backup first, then reload.</p>
          <div className="crash-actions">
            <button className="btn primary lg" onClick={this.backup}>Download backup</button>
            <button className="btn outline lg" onClick={() => location.reload()}>Reload</button>
          </div>
          <details>
            <summary>Technical details</summary>
            <pre>{String(this.state.error?.message ?? this.state.error)}</pre>
          </details>
        </div>
      </div>
    );
  }
}

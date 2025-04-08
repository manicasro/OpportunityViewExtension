export class PollingService {
  private intervalId: number | null = null;

  /**
   * Starts a polling interval.
   * @param callback The function to execute on each interval.
   * @param interval The interval duration in milliseconds.
   */
  public startPolling(callback: () => void, interval: number): void {
    this.stopPolling(); // Ensure no duplicate intervals
    this.intervalId = setInterval(callback, interval);
  }

  /**
   * Stops the polling interval.
   */
  public stopPolling(): void {
    if (this.intervalId) {
      clearInterval(this.intervalId);
      this.intervalId = null;
    }
  }
}
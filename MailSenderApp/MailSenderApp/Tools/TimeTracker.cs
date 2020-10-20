namespace MailSenderApp.Tools
{
    #region Namespaces.
    using System;
    using System.Collections.Generic;
    #endregion

    public class TimeTracker
    {
        public TimeTracker()
        {
            this.checkpoints = new Dictionary<string, TimeSpan>();
        }

        public bool AddCheckpoint(string checkpointName)
        {
            if (this.checkpoints.ContainsKey(checkpointName))
            {
                return false;
            }
            else
            {
                this.checkpoints.Add(checkpointName, new TimeSpan(DateTime.Now.Ticks));
                return true;
            }
        }

        public double GetInterval(string ckptName1, string ckptName2)
        {
            if (!this.checkpoints.ContainsKey(ckptName1) ||
                !this.checkpoints.ContainsKey(ckptName2))
            {
                throw new KeyNotFoundException(String.Format("The checkpoint name {0} or {1} are not found!", ckptName1, ckptName2));
            }

            return Convert.ToDouble(Math.Abs((this.checkpoints[ckptName2] - this.checkpoints[ckptName1]).Ticks)) / TimeSpan.TicksPerSecond;
        }

        private Dictionary<string, TimeSpan> checkpoints;
    }
}

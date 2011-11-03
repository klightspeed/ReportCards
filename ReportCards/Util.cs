using System;

namespace SouthernCluster.Util
{
    public class EventArgs<T> : EventArgs
    {
        private T val;

        public EventArgs(T val)
        {
            this.val = val;
        }

        public T Value
        {
            get
            {
                return val;
            }
        }
    }
}
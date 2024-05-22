using System;
using System.Collections.Generic;
using System.Timers;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

public class NumberServer : ExcelRtdServer
{
    private readonly List<Topic> _topics = new List<Topic>();    
    private Timer _timer;
    private int _currentValue = 0;

    public NumberServer()
    {
        _timer = new Timer();
        _timer.Interval = 1000; // 1 second interval
        _timer.Elapsed += TimerElapsed;
        _timer.Start();
    }

    private void TimerElapsed(object sender, ElapsedEventArgs e)
    {
        foreach (Topic topic in _topics)
        {
            topic.UpdateValue(_currentValue);
        }
        _currentValue++;
    }
           
    protected override void ServerTerminate()
    {
        _timer.Dispose();
        _timer = null;
    }

    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        _topics.Add(topic);
        return _currentValue;
    }

    protected override void DisconnectData(Topic topic)
    {
        _topics.Remove(topic);
        if (_topics.Count == 0)
        {
            _timer.Stop();
        }
    }
}

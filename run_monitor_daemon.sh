#!/bin/bash
# Daemon script for directory monitor
cd /root/srv/local/git

# Function to check if monitor is running
is_monitor_running() {
    pgrep -f "python3 directory_monitor.py" > /dev/null
    return $?
}

# Function to start monitor
start_monitor() {
    echo "Starting directory monitor..."
    nohup python3 directory_monitor.py >> monitor_output.log 2>&1 &
    echo "Monitor started with PID: $!"
    echo "Output being written to: monitor_output.log"
    echo "Logs are also in: logs/directory_monitor.log"
}

# Function to stop monitor
stop_monitor() {
    echo "Stopping directory monitor..."
    pkill -f "python3 directory_monitor.py"
    sleep 2
    if is_monitor_running; then
        echo "Force killing..."
        pkill -9 -f "python3 directory_monitor.py"
    fi
    echo "Monitor stopped."
}

# Main script
case "$1" in
    start)
        if is_monitor_running; then
            echo "Monitor is already running."
        else
            start_monitor
        fi
        ;;
    stop)
        stop_monitor
        ;;
    restart)
        stop_monitor
        sleep 3
        start_monitor
        ;;
    status)
        if is_monitor_running; then
            echo "Monitor is RUNNING."
            ps aux | grep directory_monitor | grep -v grep
        else
            echo "Monitor is STOPPED."
        fi
        ;;
    logs)
        echo "=== Monitor Output Log ==="
        tail -50 monitor_output.log
        echo ""
        echo "=== Directory Monitor Log ==="
        tail -50 logs/directory_monitor.log
        echo ""
        echo "=== PDF Generator Log ==="
        tail -50 logs/pdf_generator.log
        ;;
    *)
        echo "Usage: $0 {start|stop|restart|status|logs}"
        exit 1
        ;;
esac

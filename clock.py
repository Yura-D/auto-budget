from apscheduler.schedulers.blocking import BlockingScheduler
from run_formulas import make_changes


sched = BlockingScheduler()


@sched.scheduled_job('cron', day=5, hour=8)
def scheduled_job():
    make_changes()


sched.start()

import queue
import threading
import time
from datetime import datetime
import logging

class Job:
    def __init__(self, file_path, job_type, priority=0):
        self.id = f"job_{int(time.time() * 1000)}"
        self.file_path = file_path
        self.job_type = job_type
        self.priority = priority
        self.status = "PENDING"
        self.added_time = datetime.now()
        self.start_time = None
        self.end_time = None
        self.result = None
        self.error = None

class QueueManager:
    def __init__(self, processor, config, logger=None):
        self.processor = processor
        self.config = config
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.file_queue = queue.PriorityQueue()
        self.jobs = {}
        self.should_run = False
        self.processing_thread = None

    def add_job(self, file_path, job_type, priority=0):
        """Add a new job to the queue"""
        job = Job(file_path, job_type, priority)
        self.jobs[job.id] = job
        self.file_queue.put((priority, job))
        self.logger.info(f"Added job {job.id} to queue: {file_path}")
        return job.id

    def start_processing(self):
        """Start processing jobs from the queue"""
        if not self.should_run:
            self.should_run = True
            self.processing_thread = threading.Thread(target=self._process_queue)
            self.processing_thread.daemon = True
            self.processing_thread.start()
            self.logger.info("Queue processing started")

    def stop_processing(self):
        """Stop processing jobs"""
        self.should_run = False
        if self.processing_thread:
            self.processing_thread.join()
        self.logger.info("Queue processing stopped")

    def _process_queue(self):
        """Process jobs from the queue"""
        while self.should_run:
            try:
                priority, job = self.file_queue.get(timeout=1)
                if job.status == "CANCELLED":
                    continue

                job.status = "PROCESSING"
                job.start_time = datetime.now()

                try:
                    result = self.processor.process_file(job.file_path)
                    job.result = result
                    job.status = "COMPLETED"
                except Exception as e:
                    job.error = str(e)
                    job.status = "ERROR"
                    self.logger.error(f"Error processing job {job.id}: {str(e)}")

                job.end_time = datetime.now()

            except queue.Empty:
                continue
            except Exception as e:
                self.logger.error(f"Error in queue processing: {str(e)}")

    def get_job(self, job_id):
        """Get a job by ID"""
        return self.jobs.get(job_id)

    def get_all_jobs(self):
        """Get all jobs"""
        return list(self.jobs.values())

    def cancel_job(self, job_id):
        """Cancel a pending job"""
        job = self.jobs.get(job_id)
        if job and job.status == "PENDING":
            job.status = "CANCELLED"
            return True
        return False

    def clear_completed_jobs(self):
        """Clear completed jobs from the jobs dict"""
        completed = []
        for job_id, job in list(self.jobs.items()):
            if job.status in ["COMPLETED", "ERROR", "CANCELLED"]:
                completed.append(job_id)
                del self.jobs[job_id]
        return len(completed)
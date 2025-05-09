import queue
import threading
import logging
import time
import os
from datetime import datetime

class JobStatus:
    """Status constants for jobs"""
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"
    CANCELED = "CANCELED"


class Job:
    """Represents a processing job"""
    
    def __init__(self, file_path, job_type, priority=1):
        """
        Initialize a job
        
        Args:
            file_path: Path to the file to process
            job_type: Type of job (e.g., 'excel', 'ezd')
            priority: Job priority (lower number = higher priority)
        """
        self.id = f"job_{int(time.time() * 1000)}_{os.path.basename(file_path)}"
        self.file_path = file_path
        self.job_type = job_type
        self.priority = priority
        self.status = JobStatus.PENDING
        self.added_time = datetime.now()
        self.start_time = None
        self.end_time = None
        self.result = None
        self.error = None
    
    def __lt__(self, other):
        """For priority queue ordering"""
        return self.priority < other.priority


class QueueManager:
    """Manages job queues and processing"""
    
    def __init__(self, processor, config_manager, logger=None):
        """
        Initialize the queue manager
        
        Args:
            processor: Processor object to handle jobs
            config_manager: Configuration manager
            logger: Logger instance
        """
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.config_manager = config_manager
        self.processor = processor
        
        # Job queue (priority queue)
        self.job_queue = queue.PriorityQueue()
        
        # File change queue 
        self.file_queue = queue.Queue()
        
        # Track all jobs
        self.jobs = {}  # job_id -> Job
        
        # Processing threads
        self.processing_threads = []
        self.should_run = False
    
    def add_job(self, file_path, job_type, priority=1):
        """
        Add a job to the queue
        
        Args:
            file_path: Path to the file
            job_type: Type of job
            priority: Priority level (lower is higher priority)
            
        Returns:
            job_id: ID of the added job
        """
        job = Job(file_path, job_type, priority)
        self.jobs[job.id] = job
        self.job_queue.put((job.priority, job.id))
        self.logger.info(f"Added job {job.id} to queue: {file_path}")
        return job.id
    
    def get_job(self, job_id):
        """Get job by ID"""
        return self.jobs.get(job_id)
    
    def get_all_jobs(self):
        """Get all jobs"""
        return list(self.jobs.values())
    
    def cancel_job(self, job_id):
        """Cancel a pending job"""
        if job_id in self.jobs and self.jobs[job_id].status == JobStatus.PENDING:
            self.jobs[job_id].status = JobStatus.CANCELED
            self.logger.info(f"Canceled job {job_id}")
            return True
        else:
            self.logger.warning(f"Cannot cancel job {job_id} - not found or not pending")
            return False
    
    def clear_completed_jobs(self, max_age_hours=24):
        """Remove completed jobs older than specified age"""
        now = datetime.now()
        removed_count = 0
        
        for job_id in list(self.jobs.keys()):
            job = self.jobs[job_id]
            if job.status in (JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELED):
                age_hours = (now - job.added_time).total_seconds() / 3600
                if age_hours > max_age_hours:
                    del self.jobs[job_id]
                    removed_count += 1
        
        self.logger.info(f"Cleared {removed_count} old completed jobs")
        return removed_count
    
    def start_processing(self):
        """Start job processing threads"""
        if self.should_run:
            self.logger.warning("Job processing already running")
            return
            
        self.should_run = True
        
        # Start file watcher thread
        file_thread = threading.Thread(target=self._process_file_queue)
        file_thread.daemon = True
        file_thread.start()
        
        # Get max concurrent processes from config
        max_threads = self.config_manager.getint('Settings', 'max_concurrent_processes', fallback=1)
        
        # Start worker threads
        for i in range(max_threads):
            thread = threading.Thread(target=self._process_jobs, args=(i,))
            thread.daemon = True
            thread.start()
            self.processing_threads.append(thread)
        
        self.logger.info(f"Started {max_threads} job processing threads")
    
    def stop_processing(self):
        """Stop job processing"""
        self.should_run = False
        self.logger.info("Stopping job processing...")
    
    def _process_file_queue(self):
        """Process file change notifications and create jobs"""
        while self.should_run:
            try:
                # Get file from queue with timeout so we can check should_run periodically
                file_path, event_type = self.file_queue.get(timeout=1.0)
                
                # Determine file type and add appropriate job
                ext = os.path.splitext(file_path)[1].lower()
                
                if ext in ('.xls', '.xlsx'):
                    job_id = self.add_job(file_path, 'excel')
                    self.logger.info(f"Created Excel job {job_id} from file change")
                elif ext == '.ezd':
                    job_id = self.add_job(file_path, 'ezd')
                    self.logger.info(f"Created EZD job {job_id} from file change")
                else:
                    self.logger.warning(f"Unrecognized file type for {file_path}")
                
                self.file_queue.task_done()
                
            except queue.Empty:
                # No files in queue, just continue
                continue
            except Exception as e:
                self.logger.error(f"Error processing file queue: {str(e)}")
                time.sleep(1)
    
    def _process_jobs(self, worker_id):
        """Job processing worker thread"""
        self.logger.info(f"Job processor {worker_id} started")
        
        while self.should_run:
            try:
                # Get a job with timeout so we can check should_run periodically
                priority, job_id = self.job_queue.get(timeout=1.0)
                job = self.jobs.get(job_id)
                
                if not job or job.status != JobStatus.PENDING:
                    # Job was canceled or doesn't exist
                    self.job_queue.task_done()
                    continue
                
                # Process the job
                self.logger.info(f"Processing job {job_id} ({job.file_path})")
                job.status = JobStatus.RUNNING
                job.start_time = datetime.now()
                
                try:
                    if job.job_type == 'excel':
                        result = self.processor.process_excel(job.file_path)
                    elif job.job_type == 'ezd':
                        result = self.processor.process_ezd(job.file_path)
                    else:
                        raise ValueError(f"Unknown job type: {job.job_type}")
                    
                    job.result = result
                    job.status = JobStatus.COMPLETED
                    self.logger.info(f"Completed job {job_id}")
                    
                except Exception as e:
                    job.error = str(e)
                    job.status = JobStatus.FAILED
                    self.logger.error(f"Failed job {job_id}: {str(e)}")
                    
                job.end_time = datetime.now()
                self.job_queue.task_done()
                
            except queue.Empty:
                # No jobs in queue, just continue
                continue
            except Exception as e:
                self.logger.error(f"Error in job processor {worker_id}: {str(e)}")
                time.sleep(1)
        
        self.logger.info(f"Job processor {worker_id} stopped")

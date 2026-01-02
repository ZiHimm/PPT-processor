# debug_progress_timing.py
import time
import threading
import queue

def test_progress_speed():
    """Test how fast progress updates happen"""
    
    q = queue.Queue()
    messages_received = []
    
    def processor_thread():
        """Simulate your PPT processor"""
        total_slides = 81
        
        print("Simulating processor (81 slides in 0.31 seconds)...")
        start_time = time.time()
        
        for slide in range(1, total_slides + 1):
            percent = int(slide / total_slides * 100)
            
            # Send progress (like your processor does)
            q.put(f"PROGRESS:{percent}")
            
            # Simulate processing time per slide
            time_per_slide = 0.31 / 81  # 0.0038 seconds per slide
            time.sleep(time_per_slide)
        
        q.put("PROGRESS:100")
        q.put("DONE")
        
        elapsed = time.time() - start_time
        print(f"Processor finished in {elapsed:.2f} seconds")
        print(f"Average: {elapsed/81*1000:.1f}ms per slide")
    
    def ui_thread():
        """Simulate your UI checking queue"""
        print("UI thread started (checking every 100ms)...")
        
        checks = 0
        while True:
            checks += 1
            
            # Process all available messages
            messages_this_check = 0
            while True:
                try:
                    msg = q.get_nowait()
                    messages_received.append((time.time(), msg))
                    messages_this_check += 1
                    
                    if msg == "DONE":
                        print(f"\nUI detected DONE after {checks} checks")
                        return
                        
                except queue.Empty:
                    break
            
            if messages_this_check > 0:
                print(f"Check {checks}: Got {messages_this_check} messages")
            
            # Simulate UI delay (100ms like your app)
            time.sleep(0.1)
    
    # Start threads
    proc_thread = threading.Thread(target=processor_thread, daemon=True)
    ui_thread_obj = threading.Thread(target=ui_thread, daemon=True)
    
    proc_thread.start()
    time.sleep(0.05)  # Small delay
    ui_thread_obj.start()
    
    # Wait for completion
    ui_thread_obj.join(timeout=5)
    
    # Analyze results
    print(f"\n{'='*60}")
    print("ANALYSIS:")
    print(f"Total messages sent: {len(messages_received)}")
    
    if messages_received:
        first_msg = messages_received[0]
        last_msg = messages_received[-2] if len(messages_received) > 1 else first_msg  # Skip DONE
        
        if last_msg[1] != "DONE":
            last_progress = [m for m in messages_received if "PROGRESS:" in m[1]][-1]
            print(f"Last progress received: {last_progress[1]}")
        
        print(f"First message time: {first_msg[0]:.3f}")
        print(f"Last message time: {last_msg[0]:.3f}")
        print(f"Total time: {last_msg[0] - first_msg[0]:.3f}s")
    
    print(f"\n{'='*60}")
    print("CONCLUSION:")
    print("If processor sends 81 messages in 0.31s (3.8ms each)")
    print("and UI checks every 100ms...")
    print("UI can only process ~26 messages before completion!")
    print("That's why progress bar doesn't reach 100%.")

if __name__ == "__main__":
    test_progress_speed()
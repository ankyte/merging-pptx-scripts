import threading
import pickle
import os
import time
import atexit
from typing import Any, Dict, Optional, Iterable, Set


class StateManager:
    """
    A thread-safe state manager for Python applications, inspired by Streamlit's session state.
    Allows storing and accessing persistent variables across different scripts and handles concurrency.
    """
    _instance = None
    _lock = threading.RLock()
    _persist_file = "app_state.pickle"
    _auto_persist = False
    _persist_interval = 30  # seconds
    _persistence_thread = None
    
    def __new__(cls):
        """Implement singleton pattern to ensure only one state exists."""
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(StateManager, cls).__new__(cls)
                cls._instance._state = {}
                cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        """Initialize the state manager."""
        # Try to load persisted state if it exists
        self._try_load_state()
        
        # Register cleanup function to be called on exit
        atexit.register(self._cleanup)
    
    def enable_persistence(self, auto_persist: bool = True, 
                          persist_file: Optional[str] = None,
                          persist_interval: int = 30):
        """
        Enable persistence of state to disk.
        
        Args:
            auto_persist: If True, state is automatically saved periodically
            persist_file: Custom file path to save the state
            persist_interval: Interval in seconds between auto-saves
        """
        with self._lock:
            if persist_file:
                self._persist_file = persist_file
            
            self._auto_persist = auto_persist
            self._persist_interval = persist_interval
            
            # Start persistence thread if auto_persist is enabled
            if auto_persist and self._persistence_thread is None:
                self._persistence_thread = threading.Thread(
                    target=self._persistence_worker,
                    daemon=True
                )
                self._persistence_thread.start()
    
    def _persistence_worker(self):
        """Background worker that periodically saves state to disk."""
        while self._auto_persist:
            time.sleep(self._persist_interval)
            self.persist()
    
    def _try_load_state(self):
        """Try to load state from disk if it exists."""
        if os.path.exists(self._persist_file):
            try:
                with open(self._persist_file, 'rb') as f:
                    loaded_state = pickle.load(f)
                    if isinstance(loaded_state, dict):
                        self._state = loaded_state
            except (pickle.PickleError, EOFError, IOError) as e:
                print(f"Error loading persisted state: {e}")
    
    def persist(self):
        """Save state to disk."""
        with self._lock:
            try:
                with open(self._persist_file, 'wb') as f:
                    pickle.dump(self._state, f)
                return True
            except (pickle.PickleError, IOError) as e:
                print(f"Error persisting state: {e}")
                return False
    
    def _cleanup(self):
        """Clean up resources and ensure state is saved if persistence is enabled."""
        if self._auto_persist:
            self.persist()
    
    def __getitem__(self, key):
        """Allow dictionary-like access with square brackets: state['key']"""
        with self._lock:
            return self._state.get(key, None)
    
    def __setitem__(self, key, value):
        """Allow dictionary-like setting with square brackets: state['key'] = value"""
        with self._lock:
            self._state[key] = value
    
    def __delitem__(self, key):
        """Allow dictionary-like deletion with: del state['key']"""
        with self._lock:
            if key in self._state:
                del self._state[key]
    
    def __contains__(self, key):
        """Allow 'in' operator: 'key' in state"""
        with self._lock:
            return key in self._state
    
    def get(self, key, default=None):
        """Get a value with a default if key doesn't exist"""
        with self._lock:
            return self._state.get(key, default)
    
    def set(self, key, value):
        """Set a value for a key"""
        with self._lock:
            self._state[key] = value
            return value  # Return value for convenience in chaining
    
    def setdefault(self, key, default=None):
        """Set default value if key doesn't exist and return the value"""
        with self._lock:
            if key not in self._state:
                self._state[key] = default
            return self._state[key]
    
    def update(self, **kwargs):
        """Update multiple state values at once with keyword arguments"""
        with self._lock:
            self._state.update(kwargs)
    
    def clear(self):
        """Clear all state values"""
        with self._lock:
            self._state.clear()
    
    def pop(self, key, default=None):
        """Remove a key and return its value, or default if key not found"""
        with self._lock:
            return self._state.pop(key, default)
    
    def keys(self):
        """Return all keys in the state"""
        with self._lock:
            return list(self._state.keys())
    
    def values(self):
        """Return all values in the state"""
        with self._lock:
            return list(self._state.values())
    
    def items(self):
        """Return all key-value pairs in the state"""
        with self._lock:
            return list(self._state.items())
    
    def copy(self):
        """Return a copy of the state dictionary (thread-safe)"""
        with self._lock:
            return dict(self._state)
    
    def __repr__(self):
        """String representation of the state"""
        with self._lock:
            return f"StateManager({self._state})"


# Create a global instance for easy importing
state = StateManager()


# Example usage:
if __name__ == "__main__":
    # Enable persistence (optional)
    state.enable_persistence(auto_persist=True)
    
    # Example 1: Basic usage
    state['counter'] = 0
    
    # Example 2: Thread-safe increment
    def increment_counter():
        # Safely update shared state across threads
        for _ in range(1000):
            with state._lock:  # For operations that need atomicity
                current = state['counter']
                state['counter'] = current + 1
    
    # Create multiple competing threads
    threads = []
    for _ in range(10):
        t = threading.Thread(target=increment_counter)
        threads.append(t)
        t.start()
    
    # Wait for all threads to complete
    for t in threads:
        t.join()
    
    print(f"Final counter value: {state['counter']}")  # Should be 10000
    
    # Example 3: Using convenience methods
    state.update(name="Alice", age=30)
    user_id = state.setdefault('user_id', 12345)
    print(f"User: {state['name']}, ID: {user_id}")
    
    # Example 4: Loading & saving state (happens automatically if persistence enabled)
    state.persist()  # Manual save

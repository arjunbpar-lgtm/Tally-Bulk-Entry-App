import random
import uuid
import statistics

class PatternGenerator:
    def __init__(self, config):
        self.config = config
        self.min_amt = self.config.get('min_amount', 2500)
        self.max_amt = self.config.get('max_amount', 9900)
        self.min_entries = self.config.get('min_entries_per_day', 12)
        self.max_entries = self.config.get('max_entries_per_day', 40)
        self.preset = self.config.get('preset_mode', 'Mixed Rural Collection')
        self.safe_mode = self.config.get('safe_mode', False)
        self.rounding_multiple = self.config.get('rounding_multiple', 10)

    def _get_preset_bounds(self):
        # Adjust min and max based on preset
        if self.preset == 'Small Farmer Market':
            return max(self.min_amt, 3000), min(self.max_amt, 6000)
        elif self.preset == 'Heavy Collection Day':
            return max(self.min_amt, 6000), min(self.max_amt, 9900)
        elif self.preset == 'Rainy/Low Arrival Day':
            return self.min_amt, min(self.max_amt, 5000)
        else:
            return self.min_amt, self.max_amt

    def generate_initial_pattern(self, target_total, previous_average=None, seed=None):
        import math
        
        if seed is not None:
            random.seed(seed)
            
        # Ensure target_total is an integer if we are strictly avoiding decimals
        target_total = int(round(target_total))
            
        if target_total > self.max_amt * self.max_entries:
            raise ValueError(f"Impossible: Target ({target_total}) > Max Amount ({self.max_amt}) * Max Entries ({self.max_entries})")
            
        if self.min_amt > 0 and target_total < self.min_amt * self.min_entries:
            raise ValueError(f"Impossible: Target ({target_total}) < Min Amount ({self.min_amt}) * Min Entries ({self.min_entries})")
            
        p_min, p_max = self._get_preset_bounds()
        
        if self.safe_mode:
            # Narrow the ranges to make it smoother
            mid = (p_min + p_max) / 2
            p_min = (p_min + mid) / 2
            p_max = (p_max + mid) / 2

        # Decide entry count based on target_total and average
        if previous_average and self.min_amt <= previous_average <= self.max_amt:
            target_avg = previous_average
        else:
            target_avg = random.uniform(p_min, p_max)
            
        estimated_count = int(target_total / target_avg) if target_avg > 0 else self.max_entries
        
        # Calculate strict mathematical limits
        min_required_count = math.ceil(target_total / self.max_amt)
        max_allowed_count = math.floor(target_total / self.min_amt) if self.min_amt > 0 else self.max_entries
        
        # Clamp absolute bounds
        valid_min = max(self.min_entries, min_required_count)
        valid_max = min(self.max_entries, max_allowed_count)
        
        if valid_min > valid_max:
            raise ValueError(f"Impossible bounds. Target {target_total} requires between {min_required_count} and {max_allowed_count} entries, but settings limit it to {self.min_entries}-{self.max_entries}.")
        
        # Clamp estimated count within the valid boundaries
        count = max(valid_min, min(valid_max, estimated_count))
        
        # We need `count` entries summing roughly to target_total
        # We use a clustered generation approach
        cluster_center = target_total / count
        # Clamp cluster center
        cluster_center = max(p_min, min(p_max, cluster_center))
        
        entries = []
        for _ in range(count):
            # Generate a value near the cluster center to simulate similar quantities
            variance = random.uniform(0.8, 1.2) if not self.safe_mode else random.uniform(0.9, 1.1)
            val = int(cluster_center * variance)
            # Add natural rounding based on the rounding_multiple (e.g., 10, 50, 100)
            val = int(round(val / self.rounding_multiple) * self.rounding_multiple)
            
            # Ensure within hard limits
            val = max(self.min_amt, min(self.max_amt, val))
            entries.append(val)
            
        return entries

class Balancer:
    def __init__(self, config):
        self.config = config
        self.min_amt = self.config.get('min_amount', 2500)
        self.max_amt = self.config.get('max_amount', 9900)
        self.rounding_multiple = self.config.get('rounding_multiple', 10)

    def balance(self, entries, target_total):
        if not entries:
            return []
            
        current_total = sum(entries)
        diff = target_total - current_total
        
        # Sort indices randomly so we don't always adjust the same positions
        indices = list(range(len(entries)))
        random.shuffle(indices)
        
        # Iterative balancing
        # Adjust by rounding_multiple to keep it natural
        step = self.rounding_multiple if abs(diff) >= self.rounding_multiple else 1
        
        attempts = 0
        max_attempts = 10000 # Prevent infinite loops
        
        while diff != 0 and attempts < max_attempts:
            attempts += 1
            idx = random.choice(indices)
            val = entries[idx]
            
            if diff > 0:
                # Need to increase
                increment = min(step, diff)
                if val + increment <= self.max_amt:
                    entries[idx] += increment
                    diff -= increment
            else:
                # Need to decrease
                decrement = min(step, abs(diff))
                if val - decrement >= self.min_amt:
                    entries[idx] -= decrement
                    diff += decrement
                    
        if diff != 0:
            # If we still have a diff, we might need to break limits or we just fail validation
            # For now, apply directly to first available to ensure mathematical correctness,
            # Validator will catch if it breaches limits.
            entries[0] += diff
            
        return entries

class Validator:
    def __init__(self, config):
        self.config = config
        self.min_amt = self.config.get('min_amount', 2500)
        self.max_amt = self.config.get('max_amount', 9900)

    def validate(self, entries, target_total):
        target_total = int(round(target_total))
        # Hard Validation Lock rules
        if sum(entries) != target_total:
            return False, f"Total mismatch. Expected {target_total}, got {sum(entries)}."
            
        for e in entries:
            if e > self.max_amt:
                return False, f"Value {e} exceeds max limit {self.max_amt}."
            if e < self.min_amt:
                return False, f"Value {e} is below min limit {self.min_amt}."
            if e < 0:
                return False, "Negative balance found."
                
        return True, "Valid"

    def calculate_confidence(self, entries):
        if not entries:
            return 0
            
        # Basic heuristic for "Quality"
        # - Too many exact same values lowers score
        # - Standard deviation should not be wildly high or zero
        unique_vals = len(set(entries))
        ratio = unique_vals / len(entries)
        
        score = 100
        if ratio < 0.2:
            score -= 30 # Too many duplicates
        elif ratio == 1.0:
            score -= 10 # Too random (no natural clustering)
            
        # Check standard deviation
        if len(entries) > 1:
            stdev = statistics.stdev(entries)
            mean = statistics.mean(entries)
            if stdev / mean > 0.5:
                score -= 20 # Too much variance for a clustered model
                
        return max(0, min(100, score))


def generate_entries_for_day(date_str, target_total, config, previous_average=None, seed=None):
    target_total = int(round(target_total))
    if seed is None:
        seed = random.randint(1, 999999)
        
    pattern_gen = PatternGenerator(config)
    balancer = Balancer(config)
    validator = Validator(config)
    
    entries = pattern_gen.generate_initial_pattern(target_total, previous_average, seed)
    balanced_entries = balancer.balance(entries, target_total)
    
    is_valid, msg = validator.validate(balanced_entries, target_total)
    confidence = validator.calculate_confidence(balanced_entries)
    
    if not is_valid:
        raise ValueError(f"Generation failed validation: {msg}")
        
    # Format into dicts
    ledger = config.get('debit_ledger_name', 'Local Rubber Purchase')
    voucher_prefix = config.get('voucher_prefix', 'LP')
    
    batch_id = str(uuid.uuid4())[:8]
    mode = config.get('preset_mode', 'Unknown')
    
    results = []
    for i, amt in enumerate(balanced_entries, 1):
        results.append({
            'Date': date_str,
            'Voucher No': "", # Assigned dynamically at export
            'Particulars': ledger,
            'Amount': amt,
            '_BatchID': batch_id,
            '_Seed': seed,
            '_Mode': mode,
            '_Confidence': confidence
        })
        
    return results, statistics.mean(balanced_entries) if balanced_entries else None

import json
import copy
from datetime import datetime
from apps.rubber_reconstruction.generator import generate_entries_for_day, Validator
from core.settings_manager import get_active_profile

class DayReconstruction:
    def __init__(self, date_str, original_total):
        self.date_str = date_str
        self.original_total = original_total
        self.entries = []
        self.locked = False
        self.seed = None
        self.confidence = 0
        self.day_notes = ""
        self.average_val = 0
        self.batch_id = None
        self.mode = None
        
    def to_dict(self):
        return {
            'date_str': self.date_str,
            'original_total': self.original_total,
            'entries': copy.deepcopy(self.entries),
            'locked': self.locked,
            'seed': self.seed,
            'confidence': self.confidence,
            'day_notes': self.day_notes,
            'average_val': self.average_val,
            'batch_id': self.batch_id,
            'mode': self.mode
        }
        
    @classmethod
    def from_dict(cls, data):
        dr = cls(data['date_str'], data['original_total'])
        dr.entries = copy.deepcopy(data.get('entries', []))
        dr.locked = data.get('locked', False)
        dr.seed = data.get('seed')
        dr.confidence = data.get('confidence', 0)
        dr.day_notes = data.get('day_notes', "")
        dr.average_val = data.get('average_val', 0)
        dr.batch_id = data.get('batch_id')
        dr.mode = data.get('mode')
        return dr


class Action:
    def undo(self, project):
        pass
    def redo(self, project):
        pass


class RegenerateAction(Action):
    def __init__(self, date_str, old_state_dict, new_state_dict):
        self.date_str = date_str
        self.old_state_dict = old_state_dict
        self.new_state_dict = new_state_dict
        
    def undo(self, project):
        if self.date_str in project.days:
            project.days[self.date_str] = DayReconstruction.from_dict(self.old_state_dict)
            
    def redo(self, project):
        if self.date_str in project.days:
            project.days[self.date_str] = DayReconstruction.from_dict(self.new_state_dict)


class ToggleLockAction(Action):
    def __init__(self, date_str, was_locked):
        self.date_str = date_str
        self.was_locked = was_locked
        
    def undo(self, project):
        if self.date_str in project.days:
            project.days[self.date_str].locked = self.was_locked
            
    def redo(self, project):
        if self.date_str in project.days:
            project.days[self.date_str].locked = not self.was_locked


class NoteEditAction(Action):
    def __init__(self, date_str, old_note, new_note):
        self.date_str = date_str
        self.old_note = old_note
        self.new_note = new_note
        
    def undo(self, project):
        if self.date_str in project.days:
            project.days[self.date_str].day_notes = self.old_note
            
    def redo(self, project):
        if self.date_str in project.days:
            project.days[self.date_str].day_notes = self.new_note


class Project:
    def __init__(self, config=None):
        self.days = {} # date_str -> DayReconstruction
        self.config = config or get_active_profile()
        self.undo_stack = []
        self.redo_stack = []
        self.project_notes = ""
        self.file_path = None
        self.last_generation_errors = []
        
    def load_excel_data(self, daily_totals): 
        for item in daily_totals:
            date_str = item['Date']
            total = item['Total Amount']
            if date_str not in self.days:
                self.days[date_str] = DayReconstruction(date_str, total)
                
    def _execute_action(self, action):
        self.undo_stack.append(action)
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()
                
    def generate_all(self):
        self.last_generation_errors = []
        previous_averages = []
        
        sorted_date_strs = sorted(self.days.keys(), key=lambda d: datetime.strptime(d, "%d-%m-%Y"))
        
        for date_str in sorted_date_strs:
            day_obj = self.days[date_str]
            if day_obj.locked:
                if day_obj.average_val > 0:
                    previous_averages.append(day_obj.average_val)
                continue
                
            # Soft variance control smoothing
            smooth_avg = None
            if previous_averages:
                smooth_avg = sum(previous_averages[-3:]) / len(previous_averages[-3:])
                
            try:
                entries, avg = generate_entries_for_day(
                    date_str, 
                    day_obj.original_total, 
                    self.config, 
                    previous_average=smooth_avg
                )
                
                day_obj.entries = entries
                day_obj.average_val = avg
                day_obj.seed = entries[0]['_Seed'] if entries else None
                day_obj.confidence = entries[0]['_Confidence'] if entries else 0
                day_obj.batch_id = entries[0]['_BatchID'] if entries else None
                day_obj.mode = entries[0]['_Mode'] if entries else None
                
                previous_averages.append(avg)
            except Exception as e:
                self.last_generation_errors.append(f"Date {date_str} (Target: {day_obj.original_total}): {e}")
                # Clear previous entries to indicate failure
                day_obj.entries = []
                day_obj.average_val = 0
                day_obj.confidence = 0
                day_obj.day_notes = f"Error: {str(e)}"

    def regenerate_day(self, date_str):
        if date_str not in self.days:
            return
        day_obj = self.days[date_str]
        if day_obj.locked:
            return
            
        old_state = day_obj.to_dict()
        
        entries, avg = generate_entries_for_day(
            date_str, 
            day_obj.original_total, 
            self.config,
            previous_average=day_obj.average_val 
        )
        
        day_obj.entries = entries
        day_obj.average_val = avg
        day_obj.seed = entries[0]['_Seed'] if entries else None
        day_obj.confidence = entries[0]['_Confidence'] if entries else 0
        day_obj.batch_id = entries[0]['_BatchID'] if entries else None
        day_obj.mode = entries[0]['_Mode'] if entries else None
        
        new_state = day_obj.to_dict()
        self._execute_action(RegenerateAction(date_str, old_state, new_state))
        
    def toggle_lock(self, date_str):
        if date_str in self.days:
            was_locked = self.days[date_str].locked
            self.days[date_str].locked = not was_locked
            self._execute_action(ToggleLockAction(date_str, was_locked))
            
    def update_day_note(self, date_str, note):
        if date_str in self.days:
            old_note = self.days[date_str].day_notes
            self.days[date_str].day_notes = note
            self._execute_action(NoteEditAction(date_str, old_note, note))

    def undo(self):
        if self.undo_stack:
            action = self.undo_stack.pop()
            action.undo(self)
            self.redo_stack.append(action)
            return True
        return False
        
    def redo(self):
        if self.redo_stack:
            action = self.redo_stack.pop()
            action.redo(self)
            self.undo_stack.append(action)
            return True
        return False

    def save_project(self, file_path):
        data = {
            'project_notes': self.project_notes,
            'config': self.config,
            'days': {k: v.to_dict() for k, v in self.days.items()}
        }
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        self.file_path = file_path
        
    @classmethod
    def load_project(cls, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        proj = cls(config=data.get('config'))
        proj.project_notes = data.get('project_notes', "")
        proj.file_path = file_path
        
        days_dict = data.get('days', {})
        for k, v in days_dict.items():
            proj.days[k] = DayReconstruction.from_dict(v)
            
        return proj
        
    def create_export_snapshot(self, only_frozen=True):
        """
        Creates an immutable export snapshot containing validated days.
        Returns a tuple: (batch_id, list_of_DayReconstruction_copies)
        """
        validator = Validator(self.config)
        export_days = []
        
        import copy
        from datetime import datetime
        batch_id = f"EXP{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        numbering_mode = self.config.get('numbering_mode', 'Automatic (Continuous)')
        v_prefix = self.config.get('voucher_prefix', 'LP')
        v_start = int(self.config.get('voucher_start_num', 1))
        
        narration_mode = self.config.get('narration_mode', 'Automatic (Pre-filled)')
        narration_text = self.config.get('narration_text', 'Being cash paid for local rubber purchase')
        
        v_counter = v_start
        
        sorted_date_strs = sorted(self.days.keys(), key=lambda d: datetime.strptime(d, "%d-%m-%Y"))
        
        for date_str in sorted_date_strs:
            day_obj = self.days[date_str]
            if only_frozen and not day_obj.locked:
                continue
                
            if not day_obj.entries:
                continue
                
            amounts = [e['Amount'] for e in day_obj.entries]
            is_valid, msg = validator.validate(amounts, day_obj.original_total)
            if not is_valid:
                raise ValueError(f"Cannot export: Day {date_str} fails validation: {msg}")
                
            # Create a detached copy for export
            day_copy = DayReconstruction.from_dict(day_obj.to_dict())
            
            # Apply dynamic numbering and narration to the copy's entries
            for entry_copy in day_copy.entries:
                if numbering_mode == "Automatic (Continuous)":
                    entry_copy['Voucher No'] = f"{v_prefix}{v_counter:03d}"
                    v_counter += 1
                else:
                    entry_copy['Voucher No'] = ""
                    
                if narration_mode == "Automatic (Pre-filled)":
                    entry_copy['Narration'] = narration_text
                else:
                    entry_copy['Narration'] = ""
                    
            export_days.append(day_copy)
            
        return batch_id, export_days

    def get_export_data(self):
        # Legacy flat export for Excel (can also use create_export_snapshot instead)
        validator = Validator(self.config)
        export_list = []
        
        numbering_mode = self.config.get('numbering_mode', 'Automatic (Continuous)')
        v_prefix = self.config.get('voucher_prefix', 'LP')
        v_start = int(self.config.get('voucher_start_num', 1))
        
        narration_mode = self.config.get('narration_mode', 'Automatic (Pre-filled)')
        narration_text = self.config.get('narration_text', 'Being cash paid for local rubber purchase')
        
        v_counter = v_start
        
        sorted_date_strs = sorted(self.days.keys(), key=lambda d: datetime.strptime(d, "%d-%m-%Y"))
        
        for date_str in sorted_date_strs:
            day_obj = self.days[date_str]
            if not day_obj.entries:
                continue
                
            amounts = [e['Amount'] for e in day_obj.entries]
            is_valid, msg = validator.validate(amounts, day_obj.original_total)
            if not is_valid:
                raise ValueError(f"Cannot export: Day {date_str} fails validation: {msg}")
                
            for e in day_obj.entries:
                entry_copy = copy.deepcopy(e)
                if numbering_mode == "Automatic (Continuous)":
                    entry_copy['Voucher No'] = f"{v_prefix}{v_counter:03d}"
                    v_counter += 1
                else:
                    entry_copy['Voucher No'] = ""
                    
                if narration_mode == "Automatic (Pre-filled)":
                    entry_copy['Narration'] = narration_text
                else:
                    entry_copy['Narration'] = ""
                    
                export_list.append(entry_copy)
            
        return export_list

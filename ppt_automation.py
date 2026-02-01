import os
import win32com.client
from utils import log, get_media_size, calculate_centered_rect, split_rect

class PPTAutomation:
    def __init__(self, template_path, output_path, config_labels=None):
        self.template_path = template_path
        self.output_path = output_path
        self.ppt_app = None
        self.pres = None
        self.config_labels = config_labels or {}
        
    def open(self):
        log("Opening PPT Application...")
        try:
            try:
                self.ppt_app = win32com.client.Dispatch("Kwpp.Application")
                log("Using WPS Presentation.")
            except Exception:
                self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                log("Using Microsoft PowerPoint.")
            
            try:
                self.ppt_app.Visible = True
            except:
                pass
                
            self.pres = self.ppt_app.Presentations.Open(self.template_path)
            return True
        except Exception as e:
            log(f"Error opening PPT: {e}")
            return False

    def save_and_close(self):
        log("Saving and cleaning up...")
        if os.path.exists(self.output_path):
            try:
                os.remove(self.output_path)
            except Exception as e:
                log(f"Warning: Could not remove existing output file: {e}")

        try:
            self.pres.SaveAs(self.output_path)
            log(f"Saved to {self.output_path}")
        except Exception as e:
            log(f"Error saving: {e}")
            
        try:
            self.pres.Close()
            self.ppt_app.Quit()
        except:
            pass

    def _parse_tag(self, text):
        """
        Parses tag like VID_key_W600_H400_S1.0 (Case Insensitive)
        Returns: key, directives dict
        """
        parts = text.strip().split('_')
        # parts[0] is VID or IMG (handled in scan)
        if len(parts) < 2: return None, {}
        
        # Keep key as is (preserve case just in case, though Windows is case insensitive)
        key = parts[1] 
        directives = {}
        
        # Parse directives case-insensitively
        for p in parts[2:]:
            p_upper = p.upper()
            if p_upper.startswith('W'):
                try: directives['W'] = float(p_upper[1:])
                except: pass
            elif p_upper.startswith('H'):
                try: directives['H'] = float(p_upper[1:])
                except: pass
            elif p_upper.startswith('S'):
                try: directives['S'] = float(p_upper[1:])
                except: pass
                
        return key, directives

    def scan_placeholders(self):
        """
        Scans all slides for shapes starting with VID_ or IMG_ (Case Insensitive).
        Checks TextFrame first, then Name.
        Returns a list of action dicts.
        """
        actions = []
        try:
            for i, slide in enumerate(self.pres.Slides):
                # Iterate backwards to avoid issues if we were deleting (though we aren't here)
                # But actually iterating normally is fine as we just collect info.
                for shape in slide.Shapes:
                    tag_text = ""
                    # 1. Try Text Content
                    if shape.HasTextFrame:
                        try:
                            text = shape.TextFrame.TextRange.Text.strip()
                            text_upper = text.upper()
                            if text_upper.startswith("VID_") or text_upper.startswith("IMG_"):
                                tag_text = text
                        except: pass
                    
                    # 2. Try Name if Text failed
                    if not tag_text:
                        name = shape.Name
                        name_upper = name.upper()
                        if name_upper.startswith("VID_") or name_upper.startswith("IMG_"):
                            tag_text = name
                    
                    if tag_text:
                        prefix_upper = tag_text.upper()
                        prefix = "video" if prefix_upper.startswith("VID_") else "image"
                        key, directives = self._parse_tag(tag_text)
                        
                        if not key: continue

                        # Check if it's a container (user resized it)
                        is_container = False
                        if shape.Width > 50 and shape.Height > 50: # Threshold from plan
                            is_container = True
                        
                        # If explicit W/H directives exist, they override container logic
                        if 'W' in directives or 'H' in directives or 'S' in directives:
                            is_container = False 
                        
                        actions.append({
                            "slide": slide,
                            "shape": shape, 
                            "type": prefix,
                            "key": key,
                            "directives": directives,
                            "is_container": is_container
                        })
        except Exception as e:
            log(f"Error scanning slides: {e}")
        return actions

    def _insert_media_object(self, slide, media_path, left, top, width, height, is_video=True, key=None):
        """Helper to insert media (video or image) with best-effort fallbacks"""
        if is_video:
            try:
                # Try AddMediaObject2 (LinkToFile=False, SaveWithDocument=True)
                shape = slide.Shapes.AddMediaObject2(media_path, False, True, left, top, width, height)
            except:
                # Fallback
                shape = slide.Shapes.AddMediaObject(media_path, left, top, width, height)
        else:
             shape = slide.Shapes.AddPicture(media_path, False, True, left, top, width, height)
             
        # Ensure properties are set (sometimes Add methods ignore them or defaults apply)
        shape.Left = left
        shape.Top = top
        shape.Width = width
        shape.Height = height
        
        # Tagging for Object Traceability (Append Mode)
        if key:
            try:
                shape.Name = f"SmartTag_{key}"
            except:
                pass
                
        return shape

    def _add_comparison_label(self, slide, container_rect):
        """Adds the comparison text label below the content"""
        label_text = self.config_labels.get("compare_text", "（左）方案一                                                                                （右）方案二")
        left, top, width, height = container_rect
        
        # Calculate position: just below the content
        text_top = top + height + 2
        text_height = 30
        
        try:
            textbox = slide.Shapes.AddTextbox(1, left, text_top, width, text_height)
            textbox.TextFrame.TextRange.Text = label_text
            
            # Style
            textbox.TextFrame.TextRange.Font.Size = 16
            textbox.TextFrame.TextRange.Font.Color.RGB = 0x000000
            textbox.TextFrame.TextRange.Font.Bold = False
            textbox.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # Center
            
            textbox.TextFrame.MarginTop = 0
            textbox.TextFrame.MarginBottom = 0
            textbox.Line.Visible = 0
        except Exception as e:
            log(f"Warning: Could not add label text: {e}")

    def process_media_placeholder(self, action, media_path):
        """
        Replaces a placeholder with media using Smart Sizing logic.
        """
        try:
            slide = action["slide"]
            shape = action["shape"]
            directives = action["directives"]
            is_container = action["is_container"]
            is_video = (action["type"] == "video")
            
            # Get original shape rect
            orig_left = shape.Left
            orig_top = shape.Top
            orig_width = shape.Width
            orig_height = shape.Height
            
            # Delete the placeholder shape
            shape.Delete()
            
            if not os.path.exists(media_path):
                log(f"Media not found: {media_path}")
                return

            # Determine Target Rect
            media_size = get_media_size(media_path)
            if not media_size: media_size = (100, 100) # Fallback
            
            mw, mh = media_size
            
            target_left, target_top, target_width, target_height = orig_left, orig_top, orig_width, orig_height
            
            if is_container:
                # Auto-Fit into container
                target_left, target_top, target_width, target_height = calculate_centered_rect(
                    (orig_left, orig_top, orig_width, orig_height), media_size
                )
            else:
                # Anchor Mode / Directive Mode
                if 'W' in directives:
                    target_width = directives['W']
                    target_height = target_width * (mh / mw)
                elif 'H' in directives:
                    target_height = directives['H']
                    target_width = target_height * (mw / mh)
                elif 'S' in directives:
                    scale = directives['S']
                    target_width = mw * scale
                    target_height = mh * scale
                else:
                    # Default Anchor Mode: Expand from center
                    # Use 80% of slide width as default target width
                    slide_w = self.pres.PageSetup.SlideWidth
                    target_width = slide_w * 0.8
                    target_height = target_width * (mh / mw)
                    
                # Recenter based on original center
                orig_center_x = orig_left + orig_width / 2
                orig_center_y = orig_top + orig_height / 2
                
                target_left = orig_center_x - target_width / 2
                target_top = orig_center_y - target_height / 2

            # Insert
            log(f"Inserting {os.path.basename(media_path)} at Slide {slide.SlideIndex}")
            key = action.get("key")
            self._insert_media_object(slide, media_path, target_left, target_top, target_width, target_height, is_video, key=key)
            
        except Exception as e:
            log(f"Error processing placeholder {action.get('key')}: {e}")

    def insert_comparison(self, slide, shape, media_path_a, media_path_b, gap, is_video=True):
        """
        Replaces a named shape with two media files side-by-side.
        """
        try:
            # shape passed in is the placeholder object
            rect = (shape.Left, shape.Top, shape.Width, shape.Height)
            # Try to get key from shape name/text if possible to tag the new ones?
            # Actually insert_comparison is called by generate_compare_report which scans placeholders.
            # But the 'shape' object here is the placeholder which is about to be deleted.
            # We don't easily have the key passed into this function directly in arguments,
            # but we can infer it or update signature.
            # For now, let's just update the signature to accept key if we want to tag comparison results too?
            # Comparison results usually don't need to be appended again, but it's good practice.
            # However, comparison splits into A and B. How to tag? SmartTag_key_A?
            # Let's keep it simple for now and only focus on single report tagging for append.
            
            shape.Delete()
            
            rect_a_container, rect_b_container = split_rect(rect[0], rect[1], rect[2], rect[3], gap)
            
            # Process A
            if media_path_a and os.path.exists(media_path_a):
                size_a = get_media_size(media_path_a)
                final_rect_a = calculate_centered_rect(rect_a_container, size_a) if size_a else rect_a_container
                self._insert_media_object(slide, media_path_a, *final_rect_a, is_video)
            else:
                log(f"Warning: Missing A media: {media_path_a}")

            # Process B
            if media_path_b and os.path.exists(media_path_b):
                size_b = get_media_size(media_path_b)
                final_rect_b = calculate_centered_rect(rect_b_container, size_b) if size_b else rect_b_container
                self._insert_media_object(slide, media_path_b, *final_rect_b, is_video)
            else:
                log(f"Warning: Missing B media: {media_path_b}")

            self._add_comparison_label(slide, rect)
            
        except Exception as e:
            log(f"Error in insert_comparison: {e}")

    def scan_smart_tags(self):
        """
        Scans for existing shapes with name 'SmartTag_{key}'.
        Returns list of dicts: {slide, shape, key}
        """
        found = []
        try:
            for slide in self.pres.Slides:
                for shape in slide.Shapes:
                    if shape.Name.startswith("SmartTag_"):
                        key = shape.Name.replace("SmartTag_", "")
                        found.append({
                            "slide": slide,
                            "shape": shape,
                            "key": key
                        })
        except Exception as e:
            log(f"Error scanning smart tags: {e}")
        return found

    # --- Methods for Append Report (Slide Index Based) ---
    
    def _find_largest_media(self, slide):
        """Finds the largest shape on the slide, assuming it's the main content."""
        max_area = 0
        target_shape = None
        for shape in slide.Shapes:
            # Skip Title
            if shape.Type == 14:
                try:
                    if shape.PlaceholderFormat.Type in [1, 3]: continue
                except: pass
            
            if shape.Width < 50 or shape.Height < 50: continue
            
            area = shape.Width * shape.Height
            if area > max_area:
                max_area = area
                target_shape = shape
        return target_shape

    def _get_content_area(self, slide):
        """Calculates the safe content area."""
        slide_w = self.pres.PageSetup.SlideWidth
        slide_h = self.pres.PageSetup.SlideHeight
        margin_x = 20.0
        margin_bottom = 20.0
        top_margin = 60.0
        
        try:
            if slide.Shapes.HasTitle:
                title = slide.Shapes.Title
                if title.Top < slide_h / 2:
                    top_margin = title.Top + title.Height + 10.0
        except: pass
        
        left = margin_x
        top = top_margin
        width = slide_w - (2 * margin_x)
        height = slide_h - top - margin_bottom
        
        if height < slide_h * 0.5:
            top = 60.0
            height = slide_h - top - margin_bottom
            
        return (left, top, width, height)

    def insert_comparison_with_smart_tag(self, slide, existing_shape, new_media_path, gap, is_video=True):
        """
        Locates existing content (identified by SmartTag), resizes it to left, and inserts new media to right.
        """
        try:
            log(f"Found existing content on Slide {slide.SlideIndex} (SmartTag)")
            
            # Unlock aspect ratio if possible
            try: existing_shape.LockAspectRatio = 0
            except: pass

            # Get maximized container
            container_rect = self._get_content_area(slide)
            rect_a_container, rect_b_container = split_rect(*container_rect, gap)
            
            # Fit existing shape into A (Left)
            existing_w = existing_shape.Width
            existing_h = existing_shape.Height
            final_rect_a = calculate_centered_rect(rect_a_container, (existing_w, existing_h), bias_top=True)
            
            existing_shape.Left = final_rect_a[0]
            existing_shape.Top = final_rect_a[1]
            existing_shape.Width = final_rect_a[2]
            existing_shape.Height = final_rect_a[3]
            
            # Insert New Media into B (Right)
            bottom_b = final_rect_a[1] + final_rect_a[3] # Default
            
            if os.path.exists(new_media_path):
                size_b = get_media_size(new_media_path)
                final_rect_b = calculate_centered_rect(rect_b_container, size_b, bias_top=True) if size_b else rect_b_container
                self._insert_media_object(slide, new_media_path, *final_rect_b, is_video)
                bottom_b = final_rect_b[1] + final_rect_b[3]
            
            bottom_a = final_rect_a[1] + final_rect_a[3]
            text_top = max(bottom_a, bottom_b) + 2
            
            # Label
            self._add_comparison_label(slide, (container_rect[0], text_top, container_rect[2], 30))
                
        except Exception as e:
            log(f"Error in insert_comparison_with_smart_tag on Slide {slide.SlideIndex}: {e}")

    def export_smarttag_image(self, key, output_path):
        """
        Exports the shape named 'SmartTag_{key}' to the output path.
        """
        try:
            target_name = f"SmartTag_{key}"
            for slide in self.pres.Slides:
                for shape in slide.Shapes:
                    if shape.Name == target_name:
                        # 2 = ppShapeFormatPNG
                        shape.Export(output_path, 2)
                        return True
            log(f"SmartTag_{key} not found for export.")
            return False
        except Exception as e:
            log(f"Error exporting SmartTag_{key}: {e}")
            return False

    def check_val_placeholder_exists(self, key):
        """
        Checks if a text placeholder 'VAL_{key}' exists in the presentation.
        """
        target_text = f"VAL_{key}"
        try:
            for slide in self.pres.Slides:
                for shape in slide.Shapes:
                    if shape.HasTextFrame:
                        try:
                            if target_text in shape.TextFrame.TextRange.Text:
                                return True
                        except:
                            pass
        except:
            pass
        return False

    def replace_text_placeholder(self, key, value_str):
        """
        Replaces 'VAL_{key}' with value_str in all slides.
        """
        target_text = f"VAL_{key}"
        count = 0
        try:
            for slide in self.pres.Slides:
                for shape in slide.Shapes:
                    if shape.HasTextFrame:
                        try:
                            text_range = shape.TextFrame.TextRange
                            if target_text in text_range.Text:
                                text_range.Replace(FindWhat=target_text, ReplaceWhat=value_str)
                                count += 1
                        except:
                            pass
            if count > 0:
                log(f"Replaced {count} occurrences of {target_text}")
        except Exception as e:
            log(f"Error replacing text {target_text}: {e}")

    def replace_texts(self, replacements, only_first_slide=True):
        """
        Batch replace texts in the presentation.
        replacements: dict of {target: replacement}
        only_first_slide: if True, only process the first slide.
        """
        try:
            slides_to_process = []
            if only_first_slide:
                if self.pres.Slides.Count > 0:
                    slides_to_process.append(self.pres.Slides[0])
            else:
                slides_to_process = list(self.pres.Slides)
                
            count = 0
            for slide in slides_to_process:
                for shape in slide.Shapes:
                    if shape.HasTextFrame:
                        try:
                            text_range = shape.TextFrame.TextRange
                            original_text = text_range.Text
                            
                            for target, value in replacements.items():
                                if target in original_text:
                                    # Use a loop to replace all occurrences within the text range
                                    # Note: text_range.Text changes after replacement, so we check continuously
                                    # But Replace method usually handles one instance.
                                    # A safer way with COM Replace is to call it until no change or use it once if we expect one.
                                    # Given "name-Title", one replace is usually enough, but let's be robust.
                                    
                                    # Simple approach: Try replace. 
                                    if text_range.Replace(FindWhat=target, ReplaceWhat=value):
                                        count += 1
                                        # Refresh original_text for next target check if needed?
                                        # Actually Replace modifies the object in place.
                        except Exception as inner_e:
                            pass
                            
            if count > 0:
                log(f"Replaced {count} text occurrences based on metadata.")
                
        except Exception as e:
            log(f"Error in replace_texts: {e}")

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

    def scan_placeholders(self, valid_keys):
        """
        Scans all slides for shapes starting with VID_ or IMG_ that match valid keys.
        Returns a list of dicts with action info.
        """
        actions = []
        try:
            for i, slide in enumerate(self.pres.Slides):
                # slide index is 1-based, enumerate is 0-based
                for shape in slide.Shapes:
                    name = shape.Name
                    if name.startswith("VID_"):
                        key = name.replace("VID_", "")
                        if key in valid_keys:
                            actions.append({
                                "slide": slide,
                                "shape_name": name,
                                "type": "video",
                                "key": key
                            })
                    elif name.startswith("IMG_"):
                        key = name.replace("IMG_", "")
                        if key in valid_keys:
                            actions.append({
                                "slide": slide,
                                "shape_name": name,
                                "type": "image",
                                "key": key
                            })
        except Exception as e:
            log(f"Error scanning slides: {e}")
        return actions

    def _insert_media_object(self, slide, media_path, left, top, width, height, is_video=True):
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

    def insert_video_by_name(self, slide, shape_name, media_path, is_video=True):
        """
        Replaces a named shape with a video or image, centering it.
        """
        try:
            shape = slide.Shapes(shape_name)
            rect = (shape.Left, shape.Top, shape.Width, shape.Height)
            shape.Delete()
            
            if os.path.exists(media_path):
                size = get_media_size(media_path)
                if size:
                    rect = calculate_centered_rect(rect, size)
                
                left, top, width, height = rect
                log(f"Inserting {'video' if is_video else 'image'} {media_path} at Slide {slide.SlideIndex}")
                self._insert_media_object(slide, media_path, left, top, width, height, is_video)
            else:
                log(f"Warning: Media file missing: {media_path}")
                
        except Exception as e:
            log(f"Error in insert_video_by_name for {shape_name}: {e}")

    def insert_comparison(self, slide, shape_name, media_path_a, media_path_b, gap, is_video=True):
        """
        Replaces a named shape with two media files side-by-side.
        """
        try:
            shape = slide.Shapes(shape_name)
            rect = (shape.Left, shape.Top, shape.Width, shape.Height)
            shape.Delete()
            
            rect_a_container, rect_b_container = split_rect(rect[0], rect[1], rect[2], rect[3], gap)
            
            # Process A
            if os.path.exists(media_path_a):
                size_a = get_media_size(media_path_a)
                final_rect_a = calculate_centered_rect(rect_a_container, size_a) if size_a else rect_a_container
                self._insert_media_object(slide, media_path_a, *final_rect_a, is_video)
            else:
                log(f"Warning: Missing A media: {media_path_a}")

            # Process B
            if os.path.exists(media_path_b):
                size_b = get_media_size(media_path_b)
                final_rect_b = calculate_centered_rect(rect_b_container, size_b) if size_b else rect_b_container
                self._insert_media_object(slide, media_path_b, *final_rect_b, is_video)
            else:
                log(f"Warning: Missing B media: {media_path_b}")

            # Add Label
            # Note: The original code calculated text pos based on actual content bottom.
            # Here we simplify to use the container bottom for stability, or we can improve if needed.
            # But the original code logic was a bit complex with "max(content_bottom_a, content_bottom_b)".
            # Let's try to stick to the container bottom or just use the passed rect.
            # Using the original shape rect for label positioning is safer.
            self._add_comparison_label(slide, rect)
            
        except Exception as e:
            log(f"Error in insert_comparison for {shape_name}: {e}")

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

    def insert_comparison_with_existing(self, slide_index, new_media_path, gap, is_video=True):
        """
        Locates existing content on slide, resizes it to left, and inserts new media to right.
        """
        try:
            if slide_index > self.pres.Slides.Count:
                log(f"Warning: Slide {slide_index} out of range.")
                return

            slide = self.pres.Slides(slide_index)
            existing_shape = self._find_largest_media(slide)
            
            if existing_shape:
                log(f"Found existing content on Slide {slide_index}")
                
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
                if os.path.exists(new_media_path):
                    size_b = get_media_size(new_media_path)
                    final_rect_b = calculate_centered_rect(rect_b_container, size_b, bias_top=True) if size_b else rect_b_container
                    self._insert_media_object(slide, new_media_path, *final_rect_b, is_video)
                
                # Add label (using container_rect for positioning)
                # The original code used calculated content bottoms. 
                # Here we use the bottom of the content area for simplicity, or we can try to be smart.
                # Let's use the bottom of the container_rect which is where the content ends?
                # Actually, calculate_centered_rect with bias_top puts content near top.
                # So the text should be below the content.
                # Let's approximate: use the max bottom of A and B.
                bottom_a = final_rect_a[1] + final_rect_a[3]
                # For B, we need to know where we put it.
                # If we computed final_rect_b, we know.
                # If new media missing, use container bottom?
                
                # Let's just use the logic from before:
                # But here we don't have easy access to final_rect_b if we didn't calculate it (e.g. inside insert).
                # Wait, I calculated final_rect_b above.
                if os.path.exists(new_media_path):
                     bottom_b = final_rect_b[1] + final_rect_b[3]
                else:
                     bottom_b = bottom_a
                
                text_top = max(bottom_a, bottom_b) + 2
                
                # Label
                self._add_comparison_label(slide, (container_rect[0], text_top, container_rect[2], 30))
                
            else:
                log(f"Warning: No existing content found on Slide {slide_index}.")
                
        except Exception as e:
            log(f"Error in insert_comparison_with_existing on Slide {slide_index}: {e}")


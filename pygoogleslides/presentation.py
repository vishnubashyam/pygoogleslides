class Presentation:
    def __init__(self, slides_service, presentation_id):
        self.slides_service = slides_service
        self.presentation_id = presentation_id

    def fetch(self):
        """Fetch and return the presentation's full JSON structure."""
        return self.slides_service.presentations().get(
            presentationId=self.presentation_id
        ).execute()

    def batch_update(self, requests):
        """Execute a batchUpdate request on the presentation."""
        body = {'requests': requests}
        return self.slides_service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body=body
        ).execute()

    def _build_requests_for_element(self, element, placeholder, replacement, hyperlink, is_notes=False, option_title=None, font_size=None, spacing_after=None):
        """Build batchUpdate requests for a given page element if it contains the placeholder."""
        requests = []
        if 'shape' in element and 'text' in element['shape']:
            object_id = element['objectId']
            text_content = element['shape']['text'].get('textElements', [])
            full_text = ''.join([te.get('textRun', {}).get('content', '') for te in text_content])
            
            index = full_text.find(placeholder)
            if option_title and replacement.strip():
                # Detect list formatting info from the original replacement text
                body_list_info = self._detect_list_format_info(replacement)
                # Format lists (e.g., remove numbering or bullet markers)
                formatted_body = self._format_lists(replacement)
                # Process bold formatting markers in the formatted text
                final_body, bold_ranges_body = self._process_bold_formatting(formatted_body)
                combined_text = option_title + "\n" + final_body
            else:
                list_info = self._detect_list_format_info(replacement)
                formatted_text = self._format_lists(replacement)
                final_text, bold_ranges = self._process_bold_formatting(formatted_text)
                combined_text = final_text

            if placeholder in full_text:
                if hyperlink:
                    requests.append({
                        'replaceAllText': {
                            'containsText': {
                                'text': placeholder,
                                'matchCase': True
                            },
                            'replaceText': combined_text
                        }
                    })
                    if index != -1:
                        style_request = {
                            'updateTextStyle': {
                                'objectId': object_id,
                                'textRange': {
                                    'type': 'FIXED_RANGE',
                                    'startIndex': index,
                                    'endIndex': index + len(combined_text)
                                },
                                'style': {
                                    'link': {
                                        'url': hyperlink
                                    }
                                },
                                'fields': 'link'
                            }
                        }
                        if font_size is not None:
                            style_request['updateTextStyle']['style']['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
                            style_request['updateTextStyle']['fields'] += ',fontSize'
                        requests.append(style_request)
                else:
                    requests.append({
                        'replaceAllText': {
                            'containsText': {
                                'text': placeholder,
                                'matchCase': True
                            },
                            'replaceText': combined_text
                        }
                    })
                    if font_size is not None:
                        requests.append({
                            'updateTextStyle': {
                                'objectId': object_id,
                                'textRange': {
                                    'type': 'FIXED_RANGE',
                                    'startIndex': index,
                                    'endIndex': index + len(combined_text)
                                },
                                'style': {
                                    'fontSize': {'magnitude': font_size, 'unit': 'PT'}
                                },
                                'fields': 'fontSize'
                            }
                        })

                if spacing_after is not None:
                    requests.append({
                        'updateParagraphStyle': {
                            'objectId': object_id,
                            'textRange': {
                                'type': 'FIXED_RANGE',
                                'startIndex': index,
                                'endIndex': index + len(combined_text)
                            },
                            'style': {
                                'spaceAbove': {'magnitude': spacing_after, 'unit': 'PT'}
                            },
                            'fields': 'spaceAbove'
                        }
                    })

            if index != -1:
                if option_title and replacement.strip():
                    # Bold the option title (first line)
                    style_request = {
                        'updateTextStyle': {
                            'objectId': object_id,
                            'textRange': {
                                'type': 'FIXED_RANGE',
                                'startIndex': index,
                                'endIndex': index + len(option_title)
                            },
                            'style': {
                                'bold': True
                            },
                            'fields': 'bold'
                        }
                    }
                    if font_size is not None:
                        style_request['updateTextStyle']['style']['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
                        style_request['updateTextStyle']['fields'] += ',fontSize'
                    requests.append(style_request)

                    # Apply bold styling for words marked with ** in the body
                    body_start_index = index + len(option_title) + 1
                    for r_start, r_end in bold_ranges_body:
                        style_request = {
                            'updateTextStyle': {
                                'objectId': object_id,
                                'textRange': {
                                    'type': 'FIXED_RANGE',
                                    'startIndex': body_start_index + r_start,
                                    'endIndex': body_start_index + r_end
                                },
                                'style': {
                                    'bold': True
                                },
                                'fields': 'bold'
                            }
                        }
                        if font_size is not None:
                            style_request['updateTextStyle']['style']['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
                            style_request['updateTextStyle']['fields'] += ',fontSize'
                        requests.append(style_request)

                    # Apply list styling to the body (after title and newline)
                    requests.extend(self._create_list_style_requests(object_id, final_body, body_list_info, body_start_index))
                else:
                    # Apply bold styling for words marked with ** in the entire replacement text
                    for r_start, r_end in bold_ranges:
                        style_request = {
                            'updateTextStyle': {
                                'objectId': object_id,
                                'textRange': {
                                    'type': 'FIXED_RANGE',
                                    'startIndex': index + r_start,
                                    'endIndex': index + r_end
                                },
                                'style': {
                                    'bold': True
                                },
                                'fields': 'bold'
                            }
                        }
                        if font_size is not None:
                            style_request['updateTextStyle']['style']['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
                            style_request['updateTextStyle']['fields'] += ',fontSize'
                        requests.append(style_request)

                    # Apply list styling to the entire replacement text
                    requests.extend(self._create_list_style_requests(object_id, final_text, list_info, index))
        return requests

    def _format_lists(self, text):
        """Format text with numbered lists and bullet points by removing explicit markers."""
        import re  # using regex to remove markers
        lines = text.split('\n')
        formatted_lines = []
        prev_was_numbered = False
        list_indent = 0  # Track indentation level of list
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            if not stripped:
                if not prev_was_numbered:
                    formatted_lines.append(line)
                continue
            
            # Calculate current line's indentation
            current_indent = len(line) - len(line.lstrip())
            
            if re.match(r'^\d+\.\s+', stripped):
                # Remove the leading numbers, period, and space.
                new_line = re.sub(r'^\d+\.\s+', '', stripped)
                formatted_lines.append(' ' * current_indent + new_line)
                prev_was_numbered = True
                list_indent = current_indent
            elif prev_was_numbered and current_indent >= list_indent:
                # This line is part of the previous list item
                formatted_lines.append(' ' * list_indent + stripped)
                prev_was_numbered = True
            elif re.match(r'^[-*]\s+', stripped):
                # Remove the leading bullet marker and following space.
                new_line = re.sub(r'^[-*]\s+', '', stripped)
                formatted_lines.append(' ' * current_indent + new_line)
                prev_was_numbered = False
            else:
                formatted_lines.append(line)
                prev_was_numbered = False
                
        return '\n'.join(formatted_lines)

    def _detect_list_format_info(self, text):
        """Detect list formatting in text and return a list indicating for each line 
        the type ('numbered', 'bullet', or None)."""
        import re
        lines = text.split('\n')
        info = []
        prev_was_numbered = False
        list_indent = 0
        
        for line in lines:
            stripped = line.strip()
            if not stripped:
                info.append(None)
                continue
                
            # Calculate current line's indentation
            current_indent = len(line) - len(line.lstrip())
            
            if re.match(r'^\d+\.\s+', stripped):
                info.append('numbered')
                prev_was_numbered = True
                list_indent = current_indent
            elif prev_was_numbered and current_indent >= list_indent:
                # This line is part of the previous list item
                info.append('numbered')
                prev_was_numbered = True
            elif re.match(r'^[-*]\s+', stripped):
                info.append('bullet')
                prev_was_numbered = False
            else:
                info.append(None)
                prev_was_numbered = False
        return info

    def _create_list_style_requests(self, object_id, text, list_info, start_index):
        """Create requests to apply list styles to the text based on detected list format info.
        The list_info parameter is a list (of the same length as the split lines of text) 
        indicating for each line whether it is 'numbered', 'bullet', or None.
        """
        requests = []
        lines = text.split('\n')
        current_index = start_index

        for i, line in enumerate(lines):
            line_length = len(line) + 1  # +1 for newline character
            if list_info[i] == 'numbered':
                requests.append({
                    'createParagraphBullets': {
                        'objectId': object_id,
                        'textRange': {
                            'type': 'FIXED_RANGE',
                            'startIndex': current_index,
                            'endIndex': current_index + line_length
                        },
                        'bulletPreset': 'NUMBERED_DIGIT_ALPHA_ROMAN'
                    }
                })
            elif list_info[i] == 'bullet':
                requests.append({
                    'createParagraphBullets': {
                        'objectId': object_id,
                        'textRange': {
                            'type': 'FIXED_RANGE',
                            'startIndex': current_index,
                            'endIndex': current_index + line_length
                        },
                        'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                    }
                })
            current_index += line_length

        return requests

    def replace_text(self, placeholder, replacement, in_notes=False, hyperlink=None, option_title=None, font_size=None, spacing_after=None):
        """
        Replace text occurrences within slide shapes or speaker notes.
        If 'hyperlink' is provided, the replaced text will be linked.
        If an option_title is provided (and replacement is non-empty) the inserted text will have
        its first line set to option_title and formatted to be bold.
        Handles formatting for numbered lists and bullet points in both slides and speaker notes.

        Args:
            placeholder (str): The text to replace.
            replacement (str): The replacement text.
            in_notes (bool, optional): Whether to replace in speaker notes. Defaults to False.
            hyperlink (str, optional): URL to link the text to. Defaults to None.
            option_title (str, optional): Title to be bolded. Defaults to None.
            font_size (int, optional): Font size in points. Defaults to None.
            spacing_after (float, optional): Space after paragraph in points. Defaults to None.
        """
        presentation = self.fetch()
        requests = []
        for slide in presentation.get('slides', []):
            slide_requests = []
            
            # Replace text in slide content
            for element in slide.get('pageElements', []):
                reqs = self._build_requests_for_element(
                            element, placeholder, replacement, hyperlink,
                            is_notes=False, option_title=option_title,
                            font_size=font_size, spacing_after=spacing_after)
                slide_requests.extend(reqs)
            
            # Replace text in speaker notes
            if 'slideProperties' in slide and 'notesPage' in slide['slideProperties']:
                notes_page = slide['slideProperties']['notesPage']
                for element in notes_page.get('pageElements', []):
                    if 'shape' in element and 'text' in element['shape']:
                        reqs = self._build_requests_for_element(
                                    element, placeholder, replacement, hyperlink,
                                    is_notes=True, option_title=option_title,
                                    font_size=font_size, spacing_after=spacing_after)
                        slide_requests.extend(reqs)
            
            requests.extend(slide_requests)

        if requests:
            self.batch_update(requests)

    def replace_image(self, placeholder, image_url):
        """Replace an image placeholder with the actual image (applies to slide elements)."""
        request = {
            'replaceAllShapesWithImage': {
                'imageUrl': image_url,
                'containsText': {
                    'text': placeholder,
                    'matchCase': True
                }
            }
        }
        self.batch_update([request])

    def create_slide(self, predefined_layout='BLANK', insertion_index=None, object_id=None):
        """Create a new slide with the specified layout."""
        request = {
            'createSlide': {
                'slideLayoutReference': {
                    'predefinedLayout': predefined_layout
                }
            }
        }
        if insertion_index is not None:
            request['createSlide']['insertionIndex'] = insertion_index
        if object_id is not None:
            request['createSlide']['objectId'] = object_id
        self.batch_update([request])
        return object_id

    def delete_slide(self, slide_object_id):
        """Delete a slide given its object ID."""
        request = {
            'deleteObject': {
                'objectId': slide_object_id
            }
        }
        self.batch_update([request])
    
    def update_slide_layout(self, slide_object_id, new_layout_predefined):
        """Stub for updating slide layout. Not directly supported by the API."""
        raise NotImplementedError("Updating slide layout is not directly supported by the Google Slides API.")

    def _process_bold_formatting(self, text):
        """Process text for markdown-like bold markers (**bold text**) and return the text without markers plus bold ranges.
        
        Returns:
            (final_text, bold_ranges): final_text is the text with ** markers removed;
                                       bold_ranges is a list of tuples (start_index, end_index) relative to final_text.
        """
        import re
        bold_ranges = []
        result = ""
        last_idx = 0
        # Find all occurrences of **text**
        for match in re.finditer(r"\*\*(.*?)\*\*", text):
            start, end = match.span()
            # Append text before the bold marker
            result += text[last_idx:start]
            bold_text = match.group(1)
            start_in_result = len(result)
            result += bold_text
            end_in_result = len(result)
            bold_ranges.append((start_in_result, end_in_result))
            last_idx = end
        result += text[last_idx:]
        return result, bold_ranges 
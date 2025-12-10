const STORAGE_KEYS = {
    uiWidth: 'uiWidth',
    uiHeight: 'uiHeight',
    resized: 'resized',
    scope: 'scope',
  };
  
  const OPENROUTER_KEY = 'openrouterApiKey';
  const DEFAULT_UI = { width: 360, height: 520, minWidth: 280, minHeight: 200 };
  
  let cancelRequested = false;
  
  run();
  
  async function run() {
    let width = await figma.clientStorage.getAsync(STORAGE_KEYS.uiWidth);
    if (width === undefined || width === null) width = DEFAULT_UI.width;
    let height = await figma.clientStorage.getAsync(STORAGE_KEYS.uiHeight);
    if (height === undefined || height === null) height = DEFAULT_UI.height;
  
    // IMPORTANT: enable Figma theme colors for your ui.html
    figma.showUI(__html__, {
      width,
      height,
      themeColors: true, // <- this lets ui.html use --figma-color-* and figma-light/figma-dark
    });
  
    const savedScope = await figma.clientStorage.getAsync(STORAGE_KEYS.scope);
    if (savedScope) {
      figma.ui.postMessage({ type: 'scope', message: { scope: savedScope } });
    }
  
    const key = await figma.clientStorage.getAsync(OPENROUTER_KEY);
    sendOpenRouterKeyStatus(Boolean(key));
  }
  
  figma.ui.onmessage = async (msg) => {
    switch (msg.type) {
      case 'generateSlides':
        cancelRequested = false;
        await handleGenerateSlides(msg.message);
        break;
  
      case 'cancel':
        cancelRequested = true;
        figma.ui.postMessage({
          type: 'finish',
          message: { errors: emptyErrors(), summary: { status: 'Cancelled' } },
        });
        break;
  
      case 'resize': {
        const width = Math.max(DEFAULT_UI.minWidth, Number(msg.message.width));
        const height = Math.max(DEFAULT_UI.minHeight, Number(msg.message.height));
        figma.ui.resize(width, height);
        break;
      }
  
      case 'saveSize': {
        await saveSize(msg.message.width, msg.message.height, true);
        break;
      }
  
      case 'defaultSize': {
        await saveSize(DEFAULT_UI.width, DEFAULT_UI.height, false);
        figma.ui.resize(DEFAULT_UI.width, DEFAULT_UI.height);
        break;
      }
  
      case 'saveOpenRouterKey': {
        await figma.clientStorage.setAsync(OPENROUTER_KEY, (msg.message.key || '').trim());
        sendOpenRouterKeyStatus(true);
        break;
      }
  
      case 'clearOpenRouterKey': {
        await figma.clientStorage.setAsync(OPENROUTER_KEY, '');
        sendOpenRouterKeyStatus(false);
        break;
      }
  
      case 'getOpenRouterKeyStatus': {
        const key = await figma.clientStorage.getAsync(OPENROUTER_KEY);
        sendOpenRouterKeyStatus(Boolean(key));
        break;
      }
    }
  };
  
  function sendOpenRouterKeyStatus(hasKey) {
    figma.ui.postMessage({ type: 'openrouterKeyStatus', message: { hasKey } });
  }
  
  function emptyErrors() {
    return { templates: [], fonts: [], generation: [], misc: [] };
  }

  async function detectLanguage(text) {
    if (!text || text.trim().length === 0) return 'English';

    // Simple heuristic: check for common non-English characters and words
    const lowerText = text.toLowerCase();

    // Spanish indicators
    if (/\b(el|la|los|las|un|una|es|son|está|están|qué|cómo|cuándo|dónde|por qué)\b/.test(lowerText) ||
        /[áéíóúüñ¿¡]/.test(lowerText)) {
      return 'Spanish';
    }

    // French indicators
    if (/\b(le|la|les|un|une|des|et|est|sont|que|qui|quoi|comment|quand|où|pourquoi)\b/.test(lowerText) ||
        /[àâäéèêëïîôöùûüÿç]/.test(lowerText)) {
      return 'French';
    }

    // German indicators
    if (/\b(der|die|das|den|dem|des|ein|eine|einen|ist|sind|war|waren|was|wer|wie|wann|wo|warum)\b/.test(lowerText) ||
        /[äöüß]/.test(lowerText)) {
      return 'German';
    }

    // Portuguese indicators
    if (/\b(o|a|os|as|um|uma|uns|umas|é|são|está|estão|que|quem|como|quando|onde|por que)\b/.test(lowerText) ||
        /[áâãéêíóôõúç]/.test(lowerText)) {
      return 'Portuguese';
    }

    // Italian indicators
    if (/\b(il|lo|la|i|gli|le|un|uno|una|è|sono|che|chi|come|quando|dove|perché)\b/.test(lowerText) ||
        /[àèéìíîòóùú]/.test(lowerText)) {
      return 'Italian';
    }

    // Chinese/Japanese/Korean characters
    if (/[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff\uac00-\ud7af]/.test(text)) {
      if (/[\u4e00-\u9fff]/.test(text)) return 'Chinese';
      if (/[\u3040-\u309f\u30a0-\u30ff]/.test(text)) return 'Japanese';
      if (/[\uac00-\ud7af]/.test(text)) return 'Korean';
    }

    // Arabic
    if (/[\u0600-\u06ff]/.test(text)) {
      return 'Arabic';
    }

    // Russian/Cyrillic
    if (/[\u0400-\u04ff]/.test(text)) {
      return 'Russian';
    }

    // Default to English
    return 'English';
  }
  
  async function saveSize(width, height, resized) {
    await figma.clientStorage.setAsync(STORAGE_KEYS.uiWidth, width);
    await figma.clientStorage.setAsync(STORAGE_KEYS.uiHeight, height);
    await figma.clientStorage.setAsync(STORAGE_KEYS.resized, resized);
  }
  
  function collectTemplates(scope) {
    const MAX_TEMPLATES = 50; // Limit to prevent freezing with large files

    switch (scope) {
      case 'selection':
        return figma.currentPage.selection.filter((n) => n.type === 'FRAME').slice(0, MAX_TEMPLATES);
      case 'thisPage':
        return figma.currentPage.children.filter((n) => n.type === 'FRAME').slice(0, MAX_TEMPLATES);
      case 'allPages':
      case 'entireFile':
        const allFrames = figma.root.children.flatMap((page) =>
          page.children.filter((n) => n.type === 'FRAME'),
        );
        return allFrames.slice(0, MAX_TEMPLATES);
      default:
        return [];
    }
  }
  
  function getTextNodes(node) {
    const result = [];
    let nodesVisited = 0;
    const MAX_NODES = 200; // Prevent excessive recursion

    const walk = (n) => {
      if (nodesVisited >= MAX_NODES) return;
      nodesVisited++;

      if ('characters' in n) {
        result.push(n);
      } else if ('children' in n && n.children.length > 0) {
        // Limit depth and breadth to prevent performance issues
        for (let i = 0; i < Math.min(n.children.length, 20); i++) {
          walk(n.children[i]);
        }
      }
    };

    walk(node);
    return result;
  }
  
  function isMostlyNumeric(text) {
    if (!text) return false;
    const digits = (text.match(/\d/g) || []).length;
    const len = text.length;
    return len > 0 && digits / len >= 0.5;
  }
  
  function inferSlotRole(name, index, total, textSample) {
    const lower = name.toLowerCase();
    if (
      lower.includes('page') ||
      lower.includes('#') ||
      lower.includes('num') ||
      lower.includes('number') ||
      isMostlyNumeric(textSample)
    ) {
      return 'number';
    }
    if (lower.includes('title') || lower.includes('heading')) return 'title';
    if (lower.includes('sub') || lower.includes('caption')) return 'subtitle';
    if (lower.includes('bullet') || lower.includes('list') || lower.includes('item')) return 'bullets';
    if (lower.includes('body') || lower.includes('content') || lower.includes('paragraph')) return 'body';
    if (lower.includes('note')) return 'caption';
    if (index === 0) return 'title';
    if (index === 1 && total > 2) return 'subtitle';
    if (index === total - 1) return 'body';
    return 'misc';
  }
  
  function estimateSlotWords(node) {
    const length = Math.min(Math.max(node.characters.length || 60, 20), 400);
    return Math.max(4, Math.round(length / 5));
  }
  
  function estimateSlotChars(node) {
    // Estimate based on text box dimensions if available
    const currentLength = node.characters.length || 60;
    // Use current length as baseline, with reasonable bounds
    return Math.min(Math.max(currentLength, 20), 400);
  }
  
  function getAverageCharWidth(fontSize, fontFamily) {
    // Approximate character widths for common fonts (in pixels at 1px font size)
    // These are rough averages for Latin characters
    const fontWidthRatios = {
      'Inter': 0.5,
      'Arial': 0.5,
      'Helvetica': 0.5,
      'Roboto': 0.52,
      'Open Sans': 0.53,
      'Times New Roman': 0.45,
      'Georgia': 0.48,
      'Courier': 0.6,
      'Courier New': 0.6,
      'Verdana': 0.58,
      'default': 0.52
    };
    
    const ratio = fontWidthRatios[fontFamily] || fontWidthRatios['default'];
    return fontSize * ratio;
  }
  
  function estimateTextWidth(text, fontSize, fontFamily) {
    if (!text) return 0;
    const avgCharWidth = getAverageCharWidth(fontSize, fontFamily);
    return text.length * avgCharWidth;
  }
  
  function estimateMaxCharsForWidth(width, fontSize, fontFamily, lineHeight = 1.2) {
    const avgCharWidth = getAverageCharWidth(fontSize, fontFamily);
    if (avgCharWidth <= 0) return 100;
    
    const charsPerLine = Math.floor(width / avgCharWidth);
    return Math.max(10, charsPerLine);
  }
  
  function estimateMaxCharsForBox(width, height, fontSize, fontFamily, lineHeight = 1.2) {
    if (!width || !height || !fontSize) return 100;
    
    const avgCharWidth = getAverageCharWidth(fontSize, fontFamily);
    const effectiveLineHeight = fontSize * lineHeight;
    
    if (avgCharWidth <= 0 || effectiveLineHeight <= 0) return 100;
    
    const maxLines = Math.floor(height / effectiveLineHeight);
    const charsPerLine = Math.floor(width / avgCharWidth);
    
    const totalChars = maxLines * charsPerLine;
    
    // Add 10% buffer for safety
    return Math.max(10, Math.floor(totalChars * 0.9));
  }
  
  function measureTextFit(text, slot) {
    if (!text || !slot) return { fits: true, overflow: 0 };
    
    const width = slot.width || 200;
    const height = slot.height || 50;
    const fontSize = slot.fontSize || 12;
    const fontFamily = slot.fontFamily || 'Inter';
    
    const maxChars = estimateMaxCharsForBox(width, height, fontSize, fontFamily);
    const textLength = text.length;
    
    const fits = textLength <= maxChars;
    const overflow = fits ? 0 : textLength - maxChars;
    
    return { fits, overflow, maxChars, textLength };
  }
  
  function buildSlots(textNodes) {
    return textNodes.map((node, idx) => {
      // Handle mixed font sizes (figma.mixed is a Symbol)
      const fontSize = (typeof node.fontSize === 'number') ? node.fontSize : 12;
      const fontFamily = (node.fontName && typeof node.fontName === 'object' && node.fontName.family) 
        ? node.fontName.family 
        : 'Inter';
      const fontStyle = (node.fontName && typeof node.fontName === 'object' && node.fontName.style)
        ? node.fontName.style
        : 'Regular';
      const width = node.width || 200;
      const height = node.height || 50;
      
      // Use font-aware estimation for character limit
      const fontAwareCharLimit = estimateMaxCharsForBox(width, height, fontSize, fontFamily);
      
      return {
        nodeId: node.id,
        name: node.name,
        role: inferSlotRole(node.name, idx, textNodes.length, node.characters || ''),
        estimatedChars: Math.min(fontAwareCharLimit, 400),
        estimatedWords: estimateSlotWords(node),
        originalText: node.characters || '',
        fontSize: fontSize,
        fontFamily: fontFamily,
        fontStyle: fontStyle,
        width: width,
        height: height,
      };
    });
  }
  
  function buildTemplateCatalog(scope) {
    const templates = collectTemplates(scope);
    return templates.map((node, index) => {
      const textNodes = getTextNodes(node);
      const slots = buildSlots(textNodes);
      return {
        id: `template-${index}`,
        node,
        textNodes,
        layout: guessLayout(textNodes),
        name: node.name,
        isCover: index === 0, // first frame is cover
        slots,
        hasNumberSlot: slots.some((s) => s.role === 'number'),
      };
    });
  }
  
  function guessLayout(textNodes) {
    const count = textNodes.length;
    if (count <= 1) return 'title-only';
    if (count === 2) return 'title+body';
    if (count === 3) return 'title+subtitle+body';
    if (count >= 4) return 'multi-block';
    return 'generic';
  }
  
  function scoreTemplate(template, slide) {
    if (slide.templateId && slide.templateId === template.id) return -10;
    const targetCount =
      (slide.title ? 1 : 0) +
      (slide.subtitle ? 1 : 0) +
      (slide.body ? 1 : 0) +
      (slide.bullets && slide.bullets.length ? 1 : 0);
  
    const diff = Math.abs(template.textNodes.length - targetCount);
    let bonus = 0;
    if (slide.bullets && slide.bullets.length && template.textNodes.length >= slide.bullets.length + 1) {
      bonus -= 1;
    }
    if (slide.role && slide.role.toLowerCase().includes('cover') && template.isCover) {
      bonus -= 2;
    }
    return diff + bonus;
  }
  
  function pickTemplate(catalog, slide) {
    if (!catalog.length) return null;
    const byId = slide.templateId ? catalog.find((t) => t.id === slide.templateId) : null;
    if (byId) return byId;
    let best = catalog[0];
    let bestScore = Number.POSITIVE_INFINITY;
    for (const item of catalog) {
      const score = scoreTemplate(item, slide);
      if (score < bestScore) {
        best = item;
        bestScore = score;
      }
    }
    return best;
  }
  
  async function handleGenerateSlides(payload) {
    console.clear();
    const errors = emptyErrors();
  
    const prompt =
      payload.prompt !== undefined && payload.prompt !== null ? payload.prompt.trim() : '';
    const slideCount =
      Number.isFinite(payload.slideCount) && payload.slideCount > 0
        ? Math.round(payload.slideCount)
        : 0;
    const scope = payload.scope !== undefined && payload.scope !== null ? payload.scope : 'selection';
    const groupInSection = Boolean(payload.groupInSection);
  
    await figma.clientStorage.setAsync(STORAGE_KEYS.scope, scope);
  
    if (!prompt) {
      errors.generation.push('Prompt is empty.');
      sendFinish(errors);
      return;
    }
  
    if (slideCount <= 0) {
      errors.generation.push('Slide count must be a positive number.');
      sendFinish(errors);
      return;
    }
  
    const catalog = buildTemplateCatalog(scope);
    if (catalog.length === 0) {
      errors.templates.push('No template frames found in the selected scope.');
      sendFinish(errors);
      return;
    }

    // Warn if we hit the template limit
    if (catalog.length >= 50) {
      errors.misc.push(`Found ${catalog.length} templates. Limited to 50 to prevent performance issues. Consider selecting specific templates or using a smaller scope.`);
    }

    // Send progress update
    figma.ui.postMessage({
      type: 'progress',
      message: { status: 'Analyzing templates...', progress: 10 }
    });
  
    if (cancelRequested) {
      sendFinish(errors, { status: 'Cancelled' });
      return;
    }
  
    let plannedSlides = [];
    try {
      plannedSlides = await planSlidesWithDeepSeek(prompt, slideCount, catalog);
    } catch (err) {
      errors.generation.push(`Failed to plan slides: ${err.message}`);
      sendFinish(errors);
      return;
    }
  
    if (!plannedSlides.length) {
      errors.generation.push('Planner returned no slides.');
      sendFinish(errors);
      return;
    }
  
    // Ensure first slide is cover template + has role
    if (catalog[0]) {
      plannedSlides[0].templateId = plannedSlides[0].templateId || catalog[0].id;
      if (!plannedSlides[0].role) plannedSlides[0].role = 'cover';
    }
  
    let slides = [];
    try {
      slides = await generateSlidesWithDeepSeek(prompt, plannedSlides, catalog);
    } catch (err) {
      errors.generation.push(`Failed to generate slides: ${err.message}`);
      sendFinish(errors);
      return;
    }
  
    if (!slides || !slides.length) {
      errors.generation.push('AI returned no slide content.');
      sendFinish(errors);
      return;
    }
    
    // Refinement pass: Double-check and fix any text that's still too long
    if (cancelRequested) {
      sendFinish(errors, { status: 'Cancelled' });
      return;
    }
    
    figma.ui.postMessage({
      type: 'progress',
      message: { status: 'Refining text lengths...', progress: 85 }
    });
    
    try {
      slides = await refineTextLengths(slides, plannedSlides, catalog);
    } catch (err) {
      // If refinement fails, continue with original slides
      console.warn('Text refinement failed:', err.message);
    }
  
    const created = [];
    const targetPage = figma.currentPage;
    // Calculate starting position and consistent spacing
    const firstTemplate = pickTemplate(catalog, slides[0]) || catalog[0];
    const tempClone = firstTemplate.node.clone();
    const startX = tempClone.x;
    const startY = tempClone.y;
    const slideWidth = tempClone.width;
    const spacing = 80; // Gap between slides

    // Position all slides horizontally
    for (let i = 0; i < slides.length; i++) {
      if (cancelRequested) break;

      // Progress update every 3 slides
      if (i % 3 === 0) {
        figma.ui.postMessage({
          type: 'progress',
          message: {
            status: `Creating slide ${i + 1}/${slides.length}...`,
            progress: 20 + (i / slides.length) * 60
          }
        });
        // Small delay to keep UI responsive
        await new Promise(resolve => setTimeout(resolve, 1));
      }

      const slide = slides[i];
      const template = pickTemplate(catalog, slide) || catalog[0];
      const clone = template.node.clone();

      // Set slide name
      clone.name = i === 0 ? 'cover' : `slide-${i}`;

      // Position slides horizontally with consistent spacing
      clone.x = startX + (i * (slideWidth + spacing));
      clone.y = 0;

      if (clone.parent && clone.parent !== targetPage) {
        targetPage.appendChild(clone);
      }

      const textNodes = getTextNodes(clone);
      const fontErrs = await ensureFonts(textNodes);
      if (fontErrs.length) {
        // Avoid object spread for compatibility in Figma runtime
        errors.fonts.push.apply(
          errors.fonts,
          fontErrs.map((f) => `${clone.name}: ${f}`)
        );
      }

      applySlideToTemplate(clone, template, slide, prompt);
      created.push(clone);
    }
  
    let section = null;
    if (!cancelRequested && groupInSection && created.length > 0) {
      section = createSectionForClones(created, prompt);
    }
  
    if (cancelRequested) {
      errors.misc.push('Operation cancelled.');
    }
  
    sendFinish(errors, {
      status: cancelRequested ? 'Cancelled' : 'Done',
      created: created.length,
      section: section ? section.name : undefined,
    });
  }
  
  /**
   * Derive reasonable word targets from the actual template text boxes,
   * then clamp whatever Gemini suggests so it cannot explode.
   */
  function computeWordTargetsFromTemplate(template) {
    const slotsByRole = {
      title: [],
      subtitle: [],
      body: [],
      bullets: [],
    };
  
    for (const slot of template.slots || []) {
      let role = slot.role;
      if (role === 'caption') role = 'subtitle';
      if (!['title', 'subtitle', 'body', 'bullets'].includes(role)) continue;
      const est = slot.estimatedWords || 0;
      if (est > 0) slotsByRole[role].push(est);
    }
  
    const avg = (arr) =>
      arr && arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : undefined;
    const clamp = (v, min, max) =>
      v == null ? undefined : Math.min(max, Math.max(min, v));
  
    // base from actual template content
    const titleAvg = avg(slotsByRole.title);
    const subtitleAvg = avg(slotsByRole.subtitle);
    const bodyAvg = avg(slotsByRole.body);
    const bulletsAvg = avg(slotsByRole.bullets);
  
    const targets = {};
    const title = clamp(titleAvg, 2, 10);
    const subtitle = clamp(subtitleAvg, 4, 20);
    const body = clamp(bodyAvg, 20, 80);
    const bullets = clamp(bulletsAvg, 3, 14);
  
    if (title) targets.title = Math.round(title);
    if (subtitle) targets.subtitle = Math.round(subtitle);
    if (body) targets.body = Math.round(body);
    if (bullets) targets.bullets = Math.round(bullets);
    
    // Add character-based targets
    const charTargets = {};
    for (const slot of template.slots || []) {
      let role = slot.role;
      if (role === 'caption') role = 'subtitle';
      if (!['title', 'subtitle', 'body', 'bullets'].includes(role)) continue;
      if (slot.estimatedChars > 0) {
        const existing = charTargets[role] || [];
        existing.push(slot.estimatedChars);
        charTargets[role] = existing;
      }
    }
    
    if (charTargets.title && charTargets.title.length) {
      targets.titleChars = Math.round(avg(charTargets.title));
    }
    if (charTargets.subtitle && charTargets.subtitle.length) {
      targets.subtitleChars = Math.round(avg(charTargets.subtitle));
    }
    if (charTargets.body && charTargets.body.length) {
      targets.bodyChars = Math.round(avg(charTargets.body));
    }
    if (charTargets.bullets && charTargets.bullets.length) {
      targets.bulletsChars = Math.round(avg(charTargets.bullets));
    }
  
    return targets;
  }
  
  async function planSlidesWithDeepSeek(userPrompt, slideCount, catalog) {
    const apiKey = await figma.clientStorage.getAsync(OPENROUTER_KEY);
    if (!apiKey) throw new Error('OpenRouter API key not found in client storage (openrouterApiKey).');

    const minifiedTemplateSummary = catalog.map((t) => ({
      id: t.id,
      isCover: !!t.isCover,
      hasTitle: t.slots.some((s) => s.role === 'title'),
      hasSubtitle: t.slots.some((s) => s.role === 'subtitle' || s.role === 'caption'),
      hasBody: t.slots.some((s) => s.role === 'body'),
      hasBullets: t.slots.some((s) => s.role === 'bullets'),
      hasNumber: t.slots.some((s) => s.role === 'number'),
      numberExample: (t.slots.find((s) => s.role === 'number') || {}).originalText || '',
    }));

    const shortUserPrompt = userPrompt.slice(0, 500);

    // Detect language from user prompt
    const detectedLanguage = await detectLanguage(userPrompt);
  
    const planPrompt = [
      `You are an expert presentation planner. Plan exactly ${slideCount} slides for a presentation in ${detectedLanguage}.`,
      '',
      'Context:',
      '- Templates: Available slide templates with their capabilities',
      '- UserRequest: The user\'s presentation topic and requirements',
      '',
      'Task:',
      'For EACH slide (0-indexed), you must:',
      '1. Assign a clear role (cover, overview, content, example, summary, cta, conclusion, etc.)',
      '2. Select the most appropriate templateId from the available templates',
      '',
      'Critical Rules:',
      '- First slide MUST use a template where isCover === true',
      '- Match templates to roles: use bullet templates for lists, body templates for text-heavy content',
      '- Each slide serves a clear purpose in the presentation flow',
      '- Create a logical narrative structure',
      `- All content will be in ${detectedLanguage} language`,
      '',
      'Output Format (STRICT):',
      'Return ONLY a valid JSON array with exactly ${slideCount} elements.',
      'No explanations, no markdown, no code blocks - just the JSON array.',
      '',
      'Each array element must be:',
      '{',
      '  "role": string (the slide\'s purpose),',
      '  "templateId": string (from available templates)',
      '}',
      '',
      `Available Templates: ${JSON.stringify(minifiedTemplateSummary)}`,
      '',
      `User Requirements: ${shortUserPrompt}`,
      '',
      'Remember: Output ONLY the JSON array, nothing else.',
    ].join('\n');
  
    const url = 'https://openrouter.ai/api/v1/chat/completions';
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'HTTP-Referer': 'https://figma.com',
        'X-Title': 'Figma Presentation Filler Plugin'
      },
      body: JSON.stringify({
        model: 'tngtech/deepseek-r1t2-chimera:free',
        messages: [
          { role: 'user', content: planPrompt }
        ],
        temperature: 0.35,
      }),
    });
  
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`OpenRouter plan request failed (${resp.status}): ${text}`);
    }
  
    const data = await resp.json();
    const textResponse = extractOpenRouterText(data);
    if (!textResponse) throw new Error('OpenRouter planner response missing text content.');
  
    const parsed = coercePlannerArray(textResponse);
    if (!parsed || !Array.isArray(parsed)) throw new Error('Unable to parse planner JSON.');
  
    // Post-process: attach wordTargets derived from templates and clamp
    const clampGlobal = (v, min, max) =>
      typeof v === 'number' && isFinite(v) ? Math.min(max, Math.max(min, Math.round(v))) : undefined;
  
    const result = parsed.slice(0, slideCount).map((item, idx) => {
      let template = catalog.find((t) => t.id === item.templateId) || catalog[0];
  
      // Force first slide to use a cover template if available
      if (idx === 0) {
        const cover = catalog.find((t) => t.isCover) || catalog[0];
        template = cover || template;
      }
  
      const baseTargets = computeWordTargetsFromTemplate(template);
      // Avoid object spread for compatibility in Figma runtime
      const mergedTargets = Object.assign({}, baseTargets);
  
      // If the model did return wordTargets, keep them but clamp to sane ranges
      const overrides = item.wordTargets || {};
      if (overrides.title != null) {
        mergedTargets.title = clampGlobal(overrides.title, 2, 12);
      }
      if (overrides.subtitle != null) {
        mergedTargets.subtitle = clampGlobal(overrides.subtitle, 4, 24);
      }
      if (overrides.body != null) {
        mergedTargets.body = clampGlobal(overrides.body, 20, 100);
      }
      if (overrides.bullets != null) {
        mergedTargets.bullets = clampGlobal(overrides.bullets, 3, 16);
      }
  
      return {
        role: item.role || (idx === 0 ? 'cover' : 'content'),
        templateId: template.id,
        wordTargets: mergedTargets,
      };
    });
  
    return result;
  }
  
  async function refineTextLengths(slides, plannedSlides, catalog) {
    // Identify slides with text that's too long
    const slidesToRefine = [];
    
    slides.forEach((slide, idx) => {
      const plan = plannedSlides[idx] || {};
      const template = catalog.find((t) => t.id === (slide.templateId || plan.templateId)) || null;
      
      if (!template) return;
      
      const issues = [];
      const wordTargets = plan.wordTargets || {};
      
      // Check each field
      const titleSlot = template.slots.find((s) => s.role === 'title');
      const subtitleSlot = template.slots.find((s) => s.role === 'subtitle' || s.role === 'caption');
      const bodySlot = template.slots.find((s) => s.role === 'body');
      const bulletsSlot = template.slots.find((s) => s.role === 'bullets');
      
      if (slide.title && titleSlot) {
        const targetChars = wordTargets.titleChars || titleSlot.estimatedChars;
        if (slide.title.length > targetChars * 0.95) {
          issues.push({
            field: 'title',
            currentLength: slide.title.length,
            targetChars: Math.floor(targetChars * 0.85),
            currentText: slide.title
          });
        }
      }
      
      if (slide.subtitle && subtitleSlot) {
        const targetChars = wordTargets.subtitleChars || subtitleSlot.estimatedChars;
        if (slide.subtitle.length > targetChars * 0.95) {
          issues.push({
            field: 'subtitle',
            currentLength: slide.subtitle.length,
            targetChars: Math.floor(targetChars * 0.85),
            currentText: slide.subtitle
          });
        }
      }
      
      if (slide.body && bodySlot) {
        const targetChars = wordTargets.bodyChars || bodySlot.estimatedChars;
        if (slide.body.length > targetChars * 0.95) {
          issues.push({
            field: 'body',
            currentLength: slide.body.length,
            targetChars: Math.floor(targetChars * 0.85),
            currentText: slide.body
          });
        }
      }
      
      if (slide.bullets && Array.isArray(slide.bullets) && bulletsSlot) {
        const targetChars = wordTargets.bulletsChars || bulletsSlot.estimatedChars;
        slide.bullets.forEach((bullet, bIdx) => {
          if (bullet.length > targetChars * 0.95) {
            issues.push({
              field: 'bullets',
              bulletIndex: bIdx,
              currentLength: bullet.length,
              targetChars: Math.floor(targetChars * 0.85),
              currentText: bullet
            });
          }
        });
      }
      
      if (issues.length > 0) {
        slidesToRefine.push({
          slideIndex: idx,
          issues: issues,
          slide: slide
        });
      }
    });
    
    // If no slides need refinement, return as-is
    if (slidesToRefine.length === 0) {
      console.log('✓ All text lengths are acceptable');
      return slides;
    }
    
    console.log(`Refining ${slidesToRefine.length} slides with length issues`);
    
    // Ask AI to refine only the problematic fields
    const apiKey = await figma.clientStorage.getAsync(OPENROUTER_KEY);
    if (!apiKey) {
      console.warn('No API key for refinement, using truncated text');
      return slides;
    }
    
    const refinementPrompt = [
      'You are a text editor. Your ONLY job is to shorten text to fit exact character limits.',
      '',
      'Task: Rewrite the following text fields to be SHORTER while keeping the core message.',
      '',
      'Critical Rules:',
      '- Each field has a targetChars that is the MAXIMUM allowed',
      '- Your output MUST be SHORTER than the original',
      '- Aim for 80-90% of targetChars, never exceed 95%',
      '- Keep the meaning but be more concise',
      '- Remove filler words, use shorter synonyms, simplify sentence structure',
      '- No markdown, no formatting, plain text only',
      '',
      'Input format: Array of objects with { field, targetChars, currentLength, currentText }',
      'Output format: Return ONLY a JSON array with same structure but "refinedText" instead of "currentText"',
      '',
      'Example:',
      'Input: [{ "field": "title", "targetChars": 50, "currentLength": 65, "currentText": "This is a very long title that needs to be shortened significantly" }]',
      'Output: [{ "field": "title", "targetChars": 50, "refinedText": "Long Title Needs Shortening" }]',
      '',
      'Now process these texts:',
      JSON.stringify(slidesToRefine.map(sr => ({
        slideIndex: sr.slideIndex,
        issues: sr.issues.map(issue => ({
          field: issue.field,
          bulletIndex: issue.bulletIndex,
          targetChars: issue.targetChars,
          currentLength: issue.currentLength,
          currentText: issue.currentText
        }))
      }))),
      '',
      'Return ONLY the JSON array with refinedText for each field:',
    ].join('\n');
    
    const url = 'https://openrouter.ai/api/v1/chat/completions';
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'HTTP-Referer': 'https://figma.com',
        'X-Title': 'Figma Presentation Filler Plugin'
      },
      body: JSON.stringify({
        model: 'tngtech/deepseek-r1t2-chimera:free',
        messages: [
          { role: 'user', content: refinementPrompt }
        ],
        temperature: 0.2,
      }),
    });
    
    if (!resp.ok) {
      console.warn('Refinement request failed, using truncated text');
      return slides;
    }
    
    const data = await resp.json();
    const textResponse = extractOpenRouterText(data);
    if (!textResponse) {
      console.warn('Refinement response missing text');
      return slides;
    }
    
    const refinements = tryParseJsonArray(textResponse);
    if (!refinements || !Array.isArray(refinements)) {
      console.warn('Could not parse refinement response');
      return slides;
    }
    
    // Apply refinements
    refinements.forEach(refinement => {
      const slideIdx = refinement.slideIndex;
      if (slideIdx === undefined || !slides[slideIdx]) return;
      
      refinement.issues.forEach(issue => {
        if (!issue.refinedText) return;
        
        const slide = slides[slideIdx];
        
        if (issue.field === 'title') {
          slide.title = issue.refinedText;
        } else if (issue.field === 'subtitle') {
          slide.subtitle = issue.refinedText;
        } else if (issue.field === 'body') {
          slide.body = issue.refinedText;
        } else if (issue.field === 'bullets' && issue.bulletIndex !== undefined) {
          if (slide.bullets && Array.isArray(slide.bullets)) {
            slide.bullets[issue.bulletIndex] = issue.refinedText;
          }
        }
      });
    });
    
    console.log('✓ Text refinement complete');
    return slides;
  }
  
  async function generateSlidesWithDeepSeek(userPrompt, plannedSlides, catalog) {
    const apiKey = await figma.clientStorage.getAsync(OPENROUTER_KEY);
    if (!apiKey) throw new Error('OpenRouter API key not found in client storage (openrouterApiKey).');

    const minifiedTemplateSummary = catalog.map((t) => ({
      id: t.id,
      isCover: !!t.isCover,
      hasTitle: t.slots.some((s) => s.role === 'title'),
      hasSubtitle: t.slots.some((s) => s.role === 'subtitle' || s.role === 'caption'),
      hasBody: t.slots.some((s) => s.role === 'body'),
      hasBullets: t.slots.some((s) => s.role === 'bullets'),
      hasNumber: t.slots.some((s) => s.role === 'number'),
      numberExample: (t.slots.find((s) => s.role === 'number') || {}).originalText || '',
    }));

    const minifiedPlan = plannedSlides.map((p) => ({
      templateId: p.templateId,
      role: p.role,
      wordTargets: p.wordTargets || {},
    }));

    const shortUserPrompt = userPrompt.slice(0, 500);

    // Detect language from user prompt
    const detectedLanguage = await detectLanguage(userPrompt);
  
    const generationPrompt = [
      `You are an expert presentation content writer. Generate slide content in ${detectedLanguage}.`,
      '',
      'Context:',
      '- SlidePlan: Pre-planned slides with roles, templates, and strict length targets',
      '- Templates: Available template capabilities (fields they support)',
      '- UserRequest: The presentation topic and style requirements',
      '',
      'Critical Length Requirements:',
      '- Each field has wordTargets that MUST be respected',
      '- For field f with target T: generate 0.7×T to 0.9×T words',
      '- SHORTER is better than longer - text must fit in fixed text boxes',
      '- Character limits are hard constraints - exceeding them breaks the design',
      '- Think: "What\'s the minimum needed to convey this idea clearly?"',
      '',
      'Content Generation Rules:',
      `1. Language: ALL content in ${detectedLanguage}`,
      '2. Format: Plain text only - no markdown, no emojis, no bullet symbols (•, -, *)',
      '3. Bullets: Array of strings, each respecting the per-bullet word target',
      '4. Numbers: If template hasNumber, do NOT generate content for number slots',
      '5. Precision: Match slide role - covers are brief, content slides are focused',
      '',
      'IMPORTANT - Fill ALL Fields:',
      '- Generate content for EVERY field the template supports',
      '- If template hasTitle, you MUST provide title',
      '- If template hasSubtitle, you MUST provide subtitle',
      '- If template hasBody, you MUST provide body',
      '- If template hasBullets, you MUST provide bullets array',
      '- DO NOT leave any supported field empty or undefined',
      '- Use templateId and role from SlidePlan[i] for slide i',
      '',
      'Output Format (STRICT):',
      'Return ONLY a valid JSON array.',
      'No explanations, no thinking process, no markdown blocks - just the JSON.',
      '',
      'Each array element:',
      '{',
      '  "templateId": string (from plan),',
      '  "role": string (from plan),',
      '  "title"?: string (if template supports),',
      '  "subtitle"?: string (if template supports),',
      '  "body"?: string (if template supports),',
      '  "bullets"?: string[] (if template supports)',
      '}',
      '',
      `Slide Plan with Targets: ${JSON.stringify(minifiedPlan)}`,
      '',
      `Template Capabilities: ${JSON.stringify(minifiedTemplateSummary)}`,
      '',
      `User Requirements: ${shortUserPrompt}`,
      '',
      'IMPORTANT: Be concise! Text boxes are limited. Aim for 70-90% of word targets, not 100-110%.',
      'Output ONLY the JSON array now:',
    ].join('\n');
  
    const url = 'https://openrouter.ai/api/v1/chat/completions';
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'HTTP-Referer': 'https://figma.com',
        'X-Title': 'Figma Presentation Filler Plugin'
      },
      body: JSON.stringify({
        model: 'tngtech/deepseek-r1t2-chimera:free',
        messages: [
          { role: 'user', content: generationPrompt }
        ],
        temperature: 0.35,
      }),
    });
  
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`OpenRouter request failed (${resp.status}): ${text}`);
    }
  
    const data = await resp.json();
    const textResponse = extractOpenRouterText(data);
    if (!textResponse) throw new Error('OpenRouter response missing text content.');
  
    const parsed = tryParseJsonArray(textResponse);
    if (!parsed || !Array.isArray(parsed)) throw new Error('Unable to parse slide JSON from OpenRouter response.');
  
    const slides = parsed.slice(0, plannedSlides.length);
    return enforceWordTargets(slides, plannedSlides, catalog);
  }
  
  function tryParseJsonArray(content) {
    try {
      return JSON.parse(content);
    } catch (e) {
      const match = content.match(/\[([\s\S]*)\]/m);
      if (match) {
        try {
          return JSON.parse(match[0]);
        } catch (err) {
          return null;
        }
      }
      return null;
    }
  }
  
  function coercePlannerArray(content) {
    let cleaned = (content || '').trim();
  
    // direct attempt
    let parsed = tryParseJsonArray(cleaned);
    if (parsed && Array.isArray(parsed)) return parsed;
  
    // strip code fences
    cleaned = cleaned.replace(/```[a-zA-Z]*\n?/g, '').replace(/```/g, '');
    parsed = tryParseJsonArray(cleaned);
    if (parsed && Array.isArray(parsed)) return parsed;
  
    // parse as object with slides
    try {
      const obj = JSON.parse(cleaned);
      if (Array.isArray(obj)) return obj;
      if (obj && Array.isArray(obj.slides)) return obj.slides;
    } catch (e) {
      // ignore
    }
  
    // fallback: find first bracketed array
    const bracketMatch = cleaned.match(/\[[\s\S]*\]/m);
    if (bracketMatch) {
      try {
        const arr = JSON.parse(bracketMatch[0]);
        if (Array.isArray(arr)) return arr;
      } catch (e) {
        // ignore
      }
    }
    return null;
  }
  
  function extractOpenRouterText(data) {
    if (!data || !data.choices || !data.choices.length) return '';
    const first = data.choices[0];
    if (!first || !first.message || !first.message.content) return '';
    return first.message.content;
  }
  
  function countWords(text) {
    if (!text) return 0;
    return text
      .trim()
      .split(/\s+/)
      .filter(Boolean).length;
  }
  
  function countChars(text) {
    if (!text) return 0;
    return text.length;
  }
  
  function clipToTarget(text, target, maxMultiplier = 0.95) {
    if (!text) return text;
    if (!target || target <= 0) return text;
    const words = text.trim().split(/\s+/);
    const maxWords = Math.max(1, Math.round(target * maxMultiplier));
    if (words.length <= maxWords) return text;
    return words.slice(0, maxWords).join(' ');
  }
  
  function clipToCharTarget(text, charTarget, maxMultiplier = 0.95) {
    if (!text) return text;
    if (!charTarget || charTarget <= 0) return text;
    const maxChars = Math.max(1, Math.round(charTarget * maxMultiplier));
    if (text.length <= maxChars) return text;
    
    // Clip at word boundary when possible
    const words = text.trim().split(/\s+/);
    let result = '';
    for (const word of words) {
      const testResult = result ? result + ' ' + word : word;
      if (testResult.length > maxChars) {
        break;
      }
      result = testResult;
    }
    
    // If no words fit, just truncate to char limit
    if (!result && text.length > 0) {
      result = text.substring(0, maxChars);
    }
    
    return result;
  }
  
  function truncateToFit(text, slot, wordTarget, charTarget) {
    if (!text) return text;
    
    let result = text;
    
    // Pass 1: Word-based truncation with buffer (0.9x multiplier)
    if (wordTarget && wordTarget > 0) {
      result = clipToTarget(result, wordTarget, 0.9);
    }
    
    // Pass 2: Character-based truncation with buffer (0.9x multiplier)
    if (charTarget && charTarget > 0) {
      result = clipToCharTarget(result, charTarget, 0.9);
    }
    
    // Pass 3: Font-aware width check with iterative removal
    if (slot) {
      let iterations = 0;
      const maxIterations = 10;
      
      while (iterations < maxIterations) {
        const fitCheck = measureTextFit(result, slot);
        if (fitCheck.fits) break;
        
        // Remove words iteratively until it fits
        const words = result.trim().split(/\s+/);
        if (words.length <= 1) {
          // Can't remove more words, truncate chars instead
          const targetChars = fitCheck.maxChars;
          result = result.substring(0, Math.max(10, targetChars));
          break;
        }
        
        // Remove last word
        words.pop();
        result = words.join(' ');
        iterations++;
      }
    }
    
    return result;
  }
  
  function validateTextFit(slide, template) {
    const warnings = [];
    
    if (!template) return warnings;
    
    const titleSlot = template.slots.find((s) => s.role === 'title');
    const subtitleSlot = template.slots.find((s) => s.role === 'subtitle' || s.role === 'caption');
    const bodySlot = template.slots.find((s) => s.role === 'body');
    const bulletsSlot = template.slots.find((s) => s.role === 'bullets');
    
    if (slide.title && titleSlot) {
      const fitCheck = measureTextFit(slide.title, titleSlot);
      if (!fitCheck.fits) {
        warnings.push(`Title text exceeds slot by ${fitCheck.overflow} chars`);
      }
    }
    
    if (slide.subtitle && subtitleSlot) {
      const fitCheck = measureTextFit(slide.subtitle, subtitleSlot);
      if (!fitCheck.fits) {
        warnings.push(`Subtitle text exceeds slot by ${fitCheck.overflow} chars`);
      }
    }
    
    if (slide.body && bodySlot) {
      const fitCheck = measureTextFit(slide.body, bodySlot);
      if (!fitCheck.fits) {
        warnings.push(`Body text exceeds slot by ${fitCheck.overflow} chars`);
      }
    }
    
    if (slide.bullets && Array.isArray(slide.bullets) && bulletsSlot) {
      slide.bullets.forEach((bullet, idx) => {
        const fitCheck = measureTextFit(bullet, bulletsSlot);
        if (!fitCheck.fits) {
          warnings.push(`Bullet ${idx + 1} exceeds slot by ${fitCheck.overflow} chars`);
        }
      });
    }
    
    return warnings;
  }
  
  function enforceWordTargets(slides, plannedSlides, catalog) {
    return slides.map((slide, idx) => {
      const plan = plannedSlides[idx] || {};
      const template =
        catalog.find((t) => t.id === (slide.templateId || plan.templateId)) || null;
      const wordTargets = plan.wordTargets || {};
  
      const titleSlot = template ? template.slots.find((s) => s.role === 'title') : null;
      const subtitleSlot = template
        ? template.slots.find((s) => s.role === 'subtitle' || s.role === 'caption')
        : null;
      const bodySlot = template ? template.slots.find((s) => s.role === 'body') : null;
      const bulletsSlot = template ? template.slots.find((s) => s.role === 'bullets') : null;
  
      const titleTarget =
        wordTargets.title !== undefined && wordTargets.title !== null
          ? wordTargets.title
          : titleSlot
          ? titleSlot.estimatedWords
          : undefined;
      const subtitleTarget =
        wordTargets.subtitle !== undefined && wordTargets.subtitle !== null
          ? wordTargets.subtitle
          : subtitleSlot
          ? subtitleSlot.estimatedWords
          : undefined;
      const bodyTarget =
        wordTargets.body !== undefined && wordTargets.body !== null
          ? wordTargets.body
          : bodySlot
          ? bodySlot.estimatedWords
          : undefined;
      const bulletsTarget =
        wordTargets.bullets !== undefined && wordTargets.bullets !== null
          ? wordTargets.bullets
          : bulletsSlot
          ? bulletsSlot.estimatedWords
          : undefined;
  
      const bulletSlotCount = template
        ? template.slots.filter((s) => s.role === 'bullets').length
        : undefined;
  
      const trimmed = {
        templateId: slide.templateId || plan.templateId,
        role: slide.role || plan.role,
        title: slide.title,
        subtitle: slide.subtitle,
        bullets: slide.bullets,
        body: slide.body,
        wordTargets: wordTargets,
      };
  
      // Get character targets
      const titleCharTarget = wordTargets.titleChars || (titleSlot ? titleSlot.estimatedChars : undefined);
      const subtitleCharTarget = wordTargets.subtitleChars || (subtitleSlot ? subtitleSlot.estimatedChars : undefined);
      const bodyCharTarget = wordTargets.bodyChars || (bodySlot ? bodySlot.estimatedChars : undefined);
      const bulletsCharTarget = wordTargets.bulletsChars || (bulletsSlot ? bulletsSlot.estimatedChars : undefined);
      
      // Multi-pass truncation with font-aware checking
      if (trimmed.title) {
        trimmed.title = truncateToFit(trimmed.title, titleSlot, titleTarget, titleCharTarget);
      }
      if (trimmed.subtitle) {
        trimmed.subtitle = truncateToFit(trimmed.subtitle, subtitleSlot, subtitleTarget, subtitleCharTarget);
      }
      if (trimmed.body) {
        trimmed.body = truncateToFit(trimmed.body, bodySlot, bodyTarget, bodyCharTarget);
      }
  
      if (Array.isArray(trimmed.bullets)) {
        const limitedBullets = bulletSlotCount
          ? trimmed.bullets.slice(0, bulletSlotCount)
          : trimmed.bullets;
        trimmed.bullets = limitedBullets.map((b) => {
          return truncateToFit(b, bulletsSlot, bulletsTarget, bulletsCharTarget);
        });
      }
  
      // Post-process validation
      const validationWarnings = validateTextFit(trimmed, template);
      if (validationWarnings.length > 0) {
        // Log warnings but don't fail - truncation should have handled it
        console.warn(`Slide ${idx} validation warnings:`, validationWarnings);
      }
      
      return trimmed;
    });
  }
  
  // Cache for loaded fonts to avoid duplicate loading
  const loadedFonts = new Set();

  async function ensureFonts(textNodes) {
    const errors = [];
    const fontsToLoad = new Map();

    // Collect all unique fonts first
    for (const node of textNodes) {
      try {
        const fonts = await collectFonts(node);
        for (const font of fonts) {
          const key = `${font.family}-${font.style}`;
          if (!loadedFonts.has(key)) {
            fontsToLoad.set(key, font);
          }
        }
      } catch (err) {
        errors.push(err.message);
      }
    }

    // Load fonts in batches to avoid overwhelming the system
    const fontArray = Array.from(fontsToLoad.values());
    for (let i = 0; i < fontArray.length; i += 5) { // Load 5 fonts at a time
      const batch = fontArray.slice(i, i + 5);
      await Promise.all(batch.map(font => figma.loadFontAsync(font)));
      // Mark as loaded
      batch.forEach(font => {
        const key = `${font.family}-${font.style}`;
        loadedFonts.add(key);
      });
    }

    return errors;
  }
  
  async function collectFonts(node) {
    const fonts = new Map();
    if (node.fontName === figma.mixed) {
      for (let i = 0; i < node.characters.length; i++) {
        const font = node.getRangeFontName(i, i + 1);
        fonts.set(`${font.family}-${font.style}`, font);
      }
    } else {
      fonts.set(`${node.fontName.family}-${node.fontName.style}`, node.fontName);
    }
    return Array.from(fonts.values());
  }
  
  function applySlideToTemplate(frame, template, slide, prompt) {
    const slots = template.slots || [];
    const textNodes = getTextNodes(frame);
    const mapping = buildContentMap(slide, prompt);
  
    for (let i = 0; i < textNodes.length; i++) {
      const node = textNodes[i];
      const slotMeta = slots.find((s) => s.nodeId === node.id) || slots[i] || null;
      
      // Preserve numeric placeholders (e.g., page numbers / step numbers)
      if (slotMeta && slotMeta.role === 'number') {
        continue;
      }
      
      const content = pickContentForSlot(mapping, slotMeta);
      try {
        // Always set content, even if empty string - this clears old placeholder text
        node.characters = content || '';
      } catch (err) {
        // If setting content fails, try to at least clear it
        try {
          node.characters = '';
        } catch (clearErr) {
          // ignore per-node errors
        }
      }
    }
  
    // Keep the custom name set earlier (cover, slide-1, etc.) instead of overriding it
    // const namePart = slide.title || slide.role || 'Slide';
    // frame.name = `AI ${namePart}`.slice(0, 80);
  }
  
  function buildContentMap(slide, prompt) {
    return {
      title: slide.title || null,
      subtitle: slide.subtitle || null,
      bullets: slide.bullets && slide.bullets.length ? slide.bullets.join('\n') : null,
      body: slide.body || null,
      fallback: null, // Don't use prompt as fallback - leave empty if no content
    };
  }
  
  function pickContentForSlot(mapping, slotMeta) {
    // Try to fill with appropriate content, but prefer empty over mismatched content
    if (!slotMeta) {
      // No slot metadata - try title first, then empty
      return mapping.title || '';
    }
    
    switch (slotMeta.role) {
      case 'title':
        // Only use title field, don't fallback to other content
        return mapping.title || '';
      case 'subtitle':
      case 'caption':
        // Only use subtitle field, don't fallback to other content
        return mapping.subtitle || '';
      case 'bullets':
        // Only use bullets field, don't fallback to other content
        return mapping.bullets || '';
      case 'body':
        // Body can fallback to bullets if no body content
        return mapping.body || mapping.bullets || '';
      case 'misc':
        // Misc slots try body or subtitle, but prefer empty over title
        return mapping.body || mapping.subtitle || '';
      default:
        // Unknown role - try to intelligently fill
        return mapping.body || mapping.title || '';
    }
  }
  
  function createSectionForClones(clones, prompt) {
    const section = figma.createSection();
    section.name = `AI Slides – ${prompt.slice(0, 20)}`;
  
    let minX = Number.POSITIVE_INFINITY;
    let minY = Number.POSITIVE_INFINITY;
    let maxX = Number.NEGATIVE_INFINITY;
    let maxY = Number.NEGATIVE_INFINITY;
  
    for (const node of clones) {
      minX = Math.min(minX, node.x);
      minY = Math.min(minY, node.y);
      maxX = Math.max(maxX, node.x + node.width);
      maxY = Math.max(maxY, node.y + node.height);
    }
  
    // Fit section exactly to slides without padding
    section.x = minX;
    section.y = minY;
    section.resizeWithoutConstraints(maxX - minX, maxY - minY);
  
    for (const node of clones) {
      section.appendChild(node);
    }
  
    return section;
  }
  
  function sendFinish(errors, summary) {
    figma.ui.postMessage({
      type: 'finish',
      message: {
        errors,
        summary,
      },
    });
  }
  
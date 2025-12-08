const STORAGE_KEYS = {
    uiWidth: 'uiWidth',
    uiHeight: 'uiHeight',
    resized: 'resized',
    scope: 'scope',
  };
  
  const GEMINI_KEY = 'geminiApiKey';
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
  
    const key = await figma.clientStorage.getAsync(GEMINI_KEY);
    sendGeminiKeyStatus(Boolean(key));
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
  
      case 'saveGeminiKey': {
        await figma.clientStorage.setAsync(GEMINI_KEY, (msg.message.key || '').trim());
        sendGeminiKeyStatus(true);
        break;
      }
  
      case 'clearGeminiKey': {
        await figma.clientStorage.setAsync(GEMINI_KEY, '');
        sendGeminiKeyStatus(false);
        break;
      }
  
      case 'getGeminiKeyStatus': {
        const key = await figma.clientStorage.getAsync(GEMINI_KEY);
        sendGeminiKeyStatus(Boolean(key));
        break;
      }
    }
  };
  
  function sendGeminiKeyStatus(hasKey) {
    figma.ui.postMessage({ type: 'geminiKeyStatus', message: { hasKey } });
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
  
  function buildSlots(textNodes) {
    return textNodes.map((node, idx) => ({
      nodeId: node.id,
      name: node.name,
      role: inferSlotRole(node.name, idx, textNodes.length, node.characters || ''),
      estimatedChars: Math.min(Math.max(node.characters.length || 60, 20), 400),
      estimatedWords: estimateSlotWords(node),
      originalText: node.characters || '',
    }));
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
      plannedSlides = await planSlidesWithGemini(prompt, slideCount, catalog);
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
      slides = await generateSlidesWithGemini(prompt, plannedSlides, catalog);
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
  
    return targets;
  }
  
  async function planSlidesWithGemini(userPrompt, slideCount, catalog) {
    const apiKey = await figma.clientStorage.getAsync(GEMINI_KEY);
    if (!apiKey) throw new Error('Gemini API key not found in client storage (geminiApiKey).');

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
      `Plan exactly ${slideCount} slides for a presentation in ${detectedLanguage}.`,
      '',
      'You will receive:',
      '- Templates: objects with {id, isCover, hasTitle, hasSubtitle, hasBody, hasBullets, hasNumber, numberExample}.',
      '- UserRequest: short text describing topic, audience, tone, structure hints.',
      '',
      'Your job for EACH slide i (0-based):',
      '- Choose a role (e.g. "cover","overview","content","example","summary","cta","toc","divider").',
      '- Choose one templateId from Templates (do NOT invent ids).',
      '',
      'Important rules:',
      '- First slide MUST use a template where isCover === true.',
      '- Use templates that match the role: cover templates for cover, bullets templates for agenda, etc.',
      '- Do NOT worry about word counts; the plugin will derive wordTargets from template box sizes.',
      `- Generate content in ${detectedLanguage} language.`,
      '',
      'Output (STRICT):',
      '- Return ONLY a JSON array of length slideCount.',
      '- No text before or after.',
      '- item i corresponds to slide i.',
      '- Each item: {',
      '    "role": string,',
      '    "templateId": string',
      '  }',
      '',
      `Templates: ${JSON.stringify(minifiedTemplateSummary)}`,
      `UserRequest: ${shortUserPrompt}`,
    ].join('\n');
  
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: planPrompt }] }],
        generationConfig: { temperature: 0.35 },
      }),
    });
  
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`Gemini plan request failed (${resp.status}): ${text}`);
    }
  
    const data = await resp.json();
    const textResponse = extractCandidateText(data);
    if (!textResponse) throw new Error('Gemini planner response missing text content.');
  
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
  
  async function generateSlidesWithGemini(userPrompt, plannedSlides, catalog) {
    const apiKey = await figma.clientStorage.getAsync(GEMINI_KEY);
    if (!apiKey) throw new Error('Gemini API key not found in client storage (geminiApiKey).');

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
      `You write final slide text for a planned deck in ${detectedLanguage}.`,
      '',
      'You will receive:',
      '- SlidePlan: array of {role, templateId, wordTargets}.',
      '- Templates: basic info about what each template supports, including hasNumber and numberExample (keep numbers as-is).',
      '- UserRequest: topic, audience, tone, language.',
      '',
      'For EACH slide i in SlidePlan:',
      '- Use the SAME templateId and role as in SlidePlan[i].',
      '- Generate text fields that the template supports: title?, subtitle?, body?, bullets?.',
      '- Do NOT add extra fields.',
      '- If template hasNumber === true, DO NOT change that numeric slot; keep existing numberExample content.',
      '',
      'Length rules:',
      '- For every field f in wordTargets, target = wordTargets[f].',
      '- Words(f) should be roughly within 0.8×target and 1.1×target.',
      '- A word = tokens separated by spaces; treat hyphenated words as one.',
      '- If template does NOT support a field, omit it even if a target exists.',
      '',
      'Content rules:',
      `- Generate ALL content in ${detectedLanguage} language.`,
      '- Plain text only. No Markdown, no emojis, no bullet characters ("-","•","*"), no numbering. (Template number slots stay unchanged).',
      '- bullets MUST be an array of strings; each string is one bullet.',
      '- Each bullet string should also respect the bullets wordTargets range (words per bullet).',
      '- Respect role: cover/overview should be short; content slides can be denser; summary slides are concise.',
      '- Follow language/tone from UserRequest consistently.',
      '',
      'Output (STRICT):',
      '- Return ONLY a JSON array of length SlidePlan.length.',
      '- No text before or after.',
      '- item i corresponds to slide i.',
      '- Each item: {',
      '    "templateId": string,',
      '    "role": string,',
      '    "title"?: string,',
      '    "subtitle"?: string,',
      '    "bullets"?: string[],',
      '    "body"?: string',
      '  }',
      '',
      `SlidePlan: ${JSON.stringify(minifiedPlan)}`,
      `Templates: ${JSON.stringify(minifiedTemplateSummary)}`,
      `UserRequest: ${shortUserPrompt}`,
    ].join('\n');
  
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: generationPrompt }] }],
        generationConfig: { temperature: 0.35 },
      }),
    });
  
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`Gemini request failed (${resp.status}): ${text}`);
    }
  
    const data = await resp.json();
    const textResponse = extractCandidateText(data);
    if (!textResponse) throw new Error('Gemini response missing text content.');
  
    const parsed = tryParseJsonArray(textResponse);
    if (!parsed || !Array.isArray(parsed)) throw new Error('Unable to parse slide JSON from Gemini response.');
  
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
  
  function extractCandidateText(data) {
    if (!data || !data.candidates || !data.candidates.length) return '';
    const first = data.candidates[0];
    if (!first || !first.content || !first.content.parts || !first.content.parts.length) return '';
    const part = first.content.parts[0];
    return part && part.text ? part.text : '';
  }
  
  function countWords(text) {
    if (!text) return 0;
    return text
      .trim()
      .split(/\s+/)
      .filter(Boolean).length;
  }
  
  function clipToTarget(text, target, maxMultiplier = 1.05) {
    if (!text) return text;
    if (!target || target <= 0) return text;
    const words = text.trim().split(/\s+/);
    const maxWords = Math.max(1, Math.round(target * maxMultiplier));
    if (words.length <= maxWords) return text;
    return words.slice(0, maxWords).join(' ');
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
  
      if (trimmed.title) trimmed.title = clipToTarget(trimmed.title, titleTarget, 1.05);
      if (trimmed.subtitle)
        trimmed.subtitle = clipToTarget(trimmed.subtitle, subtitleTarget, 1.05);
      if (trimmed.body) trimmed.body = clipToTarget(trimmed.body, bodyTarget, 1.05);
  
      if (Array.isArray(trimmed.bullets)) {
        const limitedBullets = bulletSlotCount
          ? trimmed.bullets.slice(0, bulletSlotCount)
          : trimmed.bullets;
        trimmed.bullets = limitedBullets.map((b) =>
          clipToTarget(b, bulletsTarget, 1.05),
        );
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
        node.characters = content;
      } catch (err) {
        // ignore per-node errors
      }
    }
  
    // Keep the custom name set earlier (cover, slide-1, etc.) instead of overriding it
    // const namePart = slide.title || slide.role || 'Slide';
    // frame.name = `AI ${namePart}`.slice(0, 80);
  }
  
  function buildContentMap(slide, prompt) {
    return {
      title: slide.title,
      subtitle: slide.subtitle,
      bullets: slide.bullets && slide.bullets.length ? slide.bullets.join('\n') : undefined,
      body: slide.body,
      fallback: prompt,
    };
  }
  
  function pickContentForSlot(mapping, slotMeta) {
    if (!slotMeta) return mapping.title || mapping.body || mapping.fallback || '';
    switch (slotMeta.role) {
      case 'title':
        return mapping.title || mapping.body || mapping.fallback || '';
      case 'subtitle':
      case 'caption':
        return mapping.subtitle || mapping.body || mapping.fallback || '';
      case 'bullets':
        return mapping.bullets || mapping.body || mapping.subtitle || mapping.fallback || '';
      case 'body':
        return (
          mapping.body || mapping.bullets || mapping.subtitle || mapping.title || mapping.fallback || ''
        );
      default:
        return mapping.body || mapping.title || mapping.fallback || '';
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
  
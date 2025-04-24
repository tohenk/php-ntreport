<?php

/*
 * The MIT License
 *
 * Copyright (c) 2025 Toha <tohenk@yahoo.com>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of
 * this software and associated documentation files (the "Software"), to deal in
 * the Software without restriction, including without limitation the rights to
 * use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
 * of the Software, and to permit persons to whom the Software is furnished to do
 * so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

namespace NTLAB\Report\Filer;

use NTLAB\Report\Session\Session;
use NTLAB\Script\Core\Script;

trait Filer
{
    /**
     * @var \NTLAB\Report\Session\Session
     */
    protected $session = null;

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * Get session object.
     *
     * @return \NTLAB\Report\Session\Session
     */
    public function getSession()
    {
        return $this->session;
    }

    /**
     * Set session object.
     *
     * @param \NTLAB\Report\Session\Session $session  The session object
     * @return \NTLAB\Report\Filer\FilerInterface
     */
    public function setSession(Session $session)
    {
        $this->session = $session;

        return $this;
    }

    /**
     * Get script object.
     *
     * @return \NTLAB\Script\Core\Script
     */
    public function getScript()
    {
        if (null === $this->script) {
            $this->script = new Script();
        }

        return $this->script;
    }

    /**
     * Set the script object.
     *
     * @param \NTLAB\Script\Core\Script $script  The script object
     * @return \NTLAB\Report\Filer\FilerInterface
     */
    public function setScript(Script $script)
    {
        $this->script = $script;

        return $this;
    }
}

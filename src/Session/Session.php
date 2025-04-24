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

namespace NTLAB\Report\Session;

/**
 * Report session.
 *
 * @author Toha <tohenk@yahoo.com>
 */
class Session
{
    public const TEMPLATE = 'template';
    public const POS = 'pos';
    public const OUT = 'out';

    /**
     * @var string
     */
    protected $workId = null;

    /**
     * @var string
     */
    protected $workDir = null;

    /**
     * @var array
     */
    protected $data = [];

    /**
     * Get working directory.
     *
     * @return string
     */
    public function getWorkDir()
    {
        return $this->workDir;
    }

    /**
     * Set working directory.
     *
     * @param string $dir  The working directory
     * @return \NTLAB\Report\Report
     */
    public function setWorkDir($dir)
    {
        if (is_dir($dir)) {
            $this->workDir = $dir;
        }

        return $this;
    }

    /**
     * Create working directory.
     *
     * @return string
     */
    public function createWorkDir()
    {
        if (!is_dir($workdir = $this->getWorkIdPath($this->getWorkId()))) {
            mkdir($workdir, 0777, true);
        }

        return $workdir;
    }

    /**
     * Get work identifier.
     *
     * @return string
     */
    public function getWorkId()
    {
        if (null === $this->workId) {
            $this->workId = substr(sha1(rand(1, 999999999)), 0, 8);
        }

        return $this->workId;
    }

    /**
     * Set work identifier.
     *
     * @param string $workId  The identifier
     * @return \NTLAB\Report\Report
     */
    public function setWorkId($workId)
    {
        $this->workId = $workId;

        return $this;
    }

    /**
     * Get temporary file name.
     *
     * @param string $filename
     * @return string
     */
    public function getTempFile($filename)
    {
        return $this->createWorkDir().DIRECTORY_SEPARATOR.'~'.basename($filename);
    }

    /**
     * Load session.
     *
     * @return \NTLAB\Report\Session\Session
     */
    public function load()
    {
        if (is_readable($filename = $this->getSessionFilename()) && empty($this->data)) {
            $this->data = json_decode(file_get_contents($filename), true);
        }

        return $this;
    }

    /**
     * Save session.
     *
     * @return \NTLAB\Report\Session\Session
     */
    public function save()
    {
        if (count($this->data)) {
            file_put_contents($this->getSessionFilename(), json_encode($this->data));
        }

        return $this;
    }

    /**
     * Read data.
     *
     * @param string $key
     * @param mixed $default
     * @return mixed
     */
    public function read($key, $default = null)
    {
        return isset($this->data[$key]) ? $this->data[$key] : $default;
    }

    /**
     * Store data.
     *
     * @param string $key
     * @param mixed $value
     * @return \NTLAB\Report\Session\Session
     */
    public function store($key, $value)
    {
        $this->data[$key] = $value;

        return $this;
    }

    /**
     * Clean files.
     *
     * @return \NTLAB\Report\Session\Session
     */
    public function clean()
    {
        foreach (array_values($this->data) as $data) {
            if (is_file($data)) {
                unlink($data);
            }
        }

        return $this;
    }

    /**
     * Get path for work identity.
     *
     * @param string $workId
     * @return string
     */
    public function getWorkIdPath($workId)
    {
        if (null === ($workDir = $this->workDir)) {
            $workDir = sys_get_temp_dir();
        }

        return $workDir.DIRECTORY_SEPARATOR.'~'.$workId;
    }

    protected function getSessionFilename()
    {
        return $this->createWorkDir().DIRECTORY_SEPARATOR.'~rpt.json';
    }
}

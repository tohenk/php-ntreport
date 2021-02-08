<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2021 Toha <tohenk@yahoo.com>
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

namespace NTLAB\Report\Query;

use Propel\Runtime\Connection\ConnectionInterface;

class Pdo
{
    /**
     * @var \PDO
     */
    protected $connection = null;

    /**
     * @var \NTLAB\Report\Query\Pdo
     */
    protected static $instance = null;

    /**
     * Get Instance.
     *
     * @return \NTLAB\Report\Query\Pdo
     */
    public static function getInstance()
    {
        if (null === static::$instance) {
            static::$instance = new self();
        }
        return static::$instance;
    }

    /**
     * Set PDO connection.
     *
     * @return PDO
     */
    public function getConnection()
    {
        return $this->connection;
    }

    /**
     * Set PDO connection.
     *
     * @param \PDO $connection  The connection
     * @return \NTLAB\Report\Query\Pdo
     */
    public function setConnection($connection)
    {
        if ($connection instanceof \PDO || $connection instanceof ConnectionInterface) {
            $this->connection = $connection;
        } else {
            throw new \InvalidArgumentException('Connection must be an instance of PDO or Propel\Runtime\Connection\ConnectionInterface.');
        }
        return $this;
    }

    protected function checkConnection()
    {
        if (null === $this->connection) {
            throw new \RuntimeException('PDO connection not assigned.');
        }
    }

    /**
     * Execute SQL.
     *
     * @param string $sql  The SQL to execute
     * @param array $parameters  SQL parameters
     * @return \NTLAB\Report\Query\PdoResult[]
     */
    public function query($sql, $parameters = [])
    {
        $result = [];
        $this->checkConnection();
        $stmt = $this->connection->prepare($sql);
        foreach ($parameters as $pname => $pvalue) {
            if (is_string($pname)) {
                if (false === strpos($sql, $pname)) {
                    continue;
                }
                $stmt->bindValue($pname, $pvalue);
            }
        }
        $stmt->execute();
        while ($row = $stmt->fetch(\PDO::FETCH_ASSOC)) {
            $result[] = new PdoResult($row);
        }
        return $result;
    }
}
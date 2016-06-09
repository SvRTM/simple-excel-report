/**
 * <pre>
 * Copyright © 2012 Artem Smirnov
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * </pre>
 */
package com.github.svrtm.xlreport;

/**
 * @author Artem.Smirnov
 */
public class ReportBuilderException extends RuntimeException {

    /**
     *
     */
    private static final long serialVersionUID = -7939979556469539962L;

    /**
     *
     */
    public ReportBuilderException() {
    }

    /**
     * @param message
     */
    public ReportBuilderException(final String message) {
        super(message);
    }

    /**
     * @param cause
     */
    public ReportBuilderException(final Throwable cause) {
        super(cause);
    }

    /**
     * @param message
     * @param cause
     */
    public ReportBuilderException(final String message, final Throwable cause) {
        super(message, cause);
    }

}

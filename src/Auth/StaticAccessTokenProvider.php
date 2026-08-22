<?php

namespace Cwssrl\Office365\Auth;

use Http\Promise\FulfilledPromise;
use Http\Promise\Promise;
use Microsoft\Kiota\Abstractions\Authentication\AccessTokenProvider;
use Microsoft\Kiota\Abstractions\Authentication\AllowedHostsValidator;

/**
 * Feeds an already-obtained OAuth2 access token (e.g. from League\OAuth2\Client's
 * authorization code exchange) to the Microsoft Graph SDK, instead of having the
 * SDK perform its own token acquisition.
 */
class StaticAccessTokenProvider implements AccessTokenProvider
{
    private string $accessToken;

    private AllowedHostsValidator $allowedHostsValidator;

    public function __construct(string $accessToken)
    {
        $this->accessToken = $accessToken;
        $this->allowedHostsValidator = new AllowedHostsValidator(['graph.microsoft.com']);
    }

    public function getAuthorizationTokenAsync(string $url, array $additionalAuthenticationContext = []): Promise
    {
        if (!$this->allowedHostsValidator->isUrlHostValid($url)) {
            return new FulfilledPromise(null);
        }

        return new FulfilledPromise($this->accessToken);
    }

    public function getAllowedHostsValidator(): AllowedHostsValidator
    {
        return $this->allowedHostsValidator;
    }
}
